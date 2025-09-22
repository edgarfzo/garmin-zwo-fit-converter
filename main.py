import datetime
import xml.etree.ElementTree as ET
import os
import glob
from fit_tool.fit_file_builder import FitFileBuilder
from fit_tool.profile.messages.file_id_message import FileIdMessage
from fit_tool.profile.messages.workout_message import WorkoutMessage
from fit_tool.profile.messages.workout_step_message import WorkoutStepMessage
from fit_tool.profile.profile_type import Sport, Intensity, WorkoutStepDuration, WorkoutStepTarget, Manufacturer, FileType, WorkoutCapabilities


class zwoToFitConverter:
    def __init__(self, ftp_watts=240, use_power_for_cycling=True, power_buffer_percent=5, use_absolute_power=True, 
                 warmup_manual_advance=True, cooldown_manual_advance=False, force_warmup_power=None):
        """
        Initialize converter
        
        Args:
            ftp_watts: Your FTP in watts (used to convert percentages to absolute power)
            use_power_for_cycling: If True, use power targets for cycling workouts
            power_buffer_percent: Buffer percentage to apply (e.g., 5 for ±5%)
            use_absolute_power: If True, use absolute watts; if False, use FTP percentages
            warmup_manual_advance: If True, warmup steps require manual lap button press to advance
            cooldown_manual_advance: If True, cooldown steps require manual lap button press to advance
            force_warmup_power: If set (e.g., 0.5), override all warmup power values with this
        """
        self.ftp_watts = ftp_watts
        self.use_power_for_cycling = use_power_for_cycling
        self.power_buffer_percent = power_buffer_percent / 100.0  # Convert to decimal
        self.use_absolute_power = use_absolute_power
        self.warmup_manual_advance = warmup_manual_advance
        self.cooldown_manual_advance = cooldown_manual_advance
        self.force_warmup_power = force_warmup_power
        
        # Mapping zwo sport types to FIT sport types
        self.sport_mapping = {
            'run': Sport.RUNNING,
            'bike': Sport.CYCLING,
            'cycling': Sport.CYCLING,
            'swim': Sport.SWIMMING,
            'other': Sport.GENERIC
        }

    def _convert_power_for_fit(self, watts):
        """
        Convert power value for FIT file format
        
        Args:
            watts: Power in watts
            
        Returns:
            Power value formatted for FIT file
        """
        if self.use_absolute_power:
            # For absolute watts: add 1000 offset as per FIT specification
            return int(watts + 1000)
        else:
            # For FTP percentage: convert watts to percentage of FTP (0-1000 range)
            ftp_percentage = (watts / self.ftp_watts) * 100
            return int(round(ftp_percentage * 10))  # Scale to 0-1000 range

    def parse_zwo_file(self, zwo_file_path):
        """Parse zwo file and extract workout information"""
        tree = ET.parse(zwo_file_path)
        root = tree.getroot()
        
        # Extract basic workout info
        name = root.find('name').text if root.find('name') is not None else 'Unnamed Workout'
        description = root.find('description').text if root.find('description') is not None else ''
        sport_type = root.find('sportType').text if root.find('sportType') is not None else 'other'
        
        # Parse workout steps
        workout_element = root.find('workout')
        steps = []
        
        if workout_element is not None:
            steps = self._parse_workout_steps(workout_element, sport_type.lower())
        
        return {
            'name': name,
            'description': description,
            'sport': sport_type.lower(),
            'steps': steps
        }

    def _parse_workout_steps(self, workout_element, sport='bike'):
        """Parse workout steps and expand intervals"""
        steps = []
        
        for element in workout_element:
            if element.tag == 'Warmup':
                step = self._parse_warmup(element, sport)
                steps.append(step)
                
            elif element.tag == 'Cooldown':
                step = self._parse_cooldown(element, sport)
                steps.append(step)
                
            elif element.tag == 'IntervalsT':
                interval_steps = self._parse_intervals(element, sport)
                steps.extend(interval_steps)
                
            elif element.tag == 'SteadyState':
                step = self._parse_steady_state(element, sport)
                steps.append(step)
        
        return steps

    def _should_use_power(self, sport):
        """Determine if we should use power targets for this sport"""
        return self.use_power_for_cycling and sport in ['bike', 'cycling']

    def _apply_power_buffer_watts(self, power_percentage):
        """Convert power percentage to absolute watts and apply buffer"""
        # Convert percentage to absolute watts
        target_watts = power_percentage * self.ftp_watts
        
        # Apply buffer
        low_watts = target_watts * (1 - self.power_buffer_percent)
        high_watts = target_watts * (1 + self.power_buffer_percent)
        
        # Round to integers
        return int(round(low_watts)), int(round(high_watts))

    def _parse_warmup(self, element, sport='bike'):
        """Parse warmup step"""
        duration = int(element.get('Duration', 600))  # Duration in seconds
        power_low = float(element.get('PowerLow', 0.60))
        power_high = float(element.get('PowerHigh', 0.70))
        
        # Override warmup power if force_warmup_power is set
        if self.force_warmup_power is not None:
            power_low = self.force_warmup_power
            power_high = self.force_warmup_power
            print(f"  Warmup power overridden to {self.force_warmup_power} ({self.force_warmup_power*100:.0f}%)")
        
        # Determine duration type based on manual advance setting
        if self.warmup_manual_advance:
            duration_type = WorkoutStepDuration.OPEN
            duration_value = 0  # Not used for open duration
            step_name = 'Warm up (Press LAP when ready)'
        else:
            duration_type = WorkoutStepDuration.TIME
            duration_value = duration * 1000
            step_name = 'Warm up'
        
        if self._should_use_power(sport):
            # Convert to absolute watts with buffer
            if power_low == power_high:
                # Single power value, apply buffer
                power_low_watts, power_high_watts = self._apply_power_buffer_watts(power_low)
            else:
                # Range provided, convert each to watts and apply individual buffers
                power_low_watts, _ = self._apply_power_buffer_watts(power_low)
                _, power_high_watts = self._apply_power_buffer_watts(power_high)
            
            # Convert to FIT format
            fit_power_low = self._convert_power_for_fit(power_low_watts)
            fit_power_high = self._convert_power_for_fit(power_high_watts)
            
            return {
                'wkt_step_name': step_name,
                'intensity': Intensity.WARMUP,
                'duration_type': duration_type,
                'duration_value': duration_value,
                'target_type': WorkoutStepTarget.POWER,
                'target_value': 0,  # Set to 0 when using custom ranges
                'custom_target_value_low': fit_power_low,
                'custom_target_value_high': fit_power_high
            }
        else:
            # Use heart rate zones - use the forced power value if set
            if self.force_warmup_power is not None:
                hr_zone = self._power_to_heart_rate_zone(self.force_warmup_power)
                print(f"  Warmup HR zone calculated from forced power {self.force_warmup_power}: Zone {hr_zone}")
            else:
                avg_power = (power_low + power_high) / 2
                hr_zone = self._power_to_heart_rate_zone(avg_power)
                print(f"  Warmup HR zone calculated from average power {avg_power}: Zone {hr_zone}")
            
            return {
                'wkt_step_name': step_name,
                'intensity': Intensity.WARMUP,
                'duration_type': duration_type,
                'duration_value': duration_value,
                'target_type': WorkoutStepTarget.HEART_RATE,
                'target_value': hr_zone
            }

    def _parse_cooldown(self, element, sport='bike'):
        """Parse cooldown step"""
        duration = int(element.get('Duration', 600))  # Duration in seconds
        power_low = float(element.get('PowerLow', 0.60))
        power_high = float(element.get('PowerHigh', 0.65))
        
        # Determine duration type based on manual advance setting
        if self.cooldown_manual_advance:
            duration_type = WorkoutStepDuration.OPEN
            duration_value = 0  # Not used for open duration
            step_name = 'Cool down (Press LAP when ready)'
        else:
            duration_type = WorkoutStepDuration.TIME
            duration_value = duration * 1000
            step_name = 'Cool down'
        
        if self._should_use_power(sport):
            # Convert to absolute watts with buffer
            if power_low == power_high:
                # Single power value, apply buffer
                power_low_watts, power_high_watts = self._apply_power_buffer_watts(power_low)
            else:
                # Range provided, convert each to watts and apply individual buffers
                power_low_watts, _ = self._apply_power_buffer_watts(power_low)
                _, power_high_watts = self._apply_power_buffer_watts(power_high)
            
            # Convert to FIT format
            fit_power_low = self._convert_power_for_fit(power_low_watts)
            fit_power_high = self._convert_power_for_fit(power_high_watts)
            
            return {
                'wkt_step_name': step_name,
                'intensity': Intensity.COOLDOWN,
                'duration_type': duration_type,
                'duration_value': duration_value,
                'target_type': WorkoutStepTarget.POWER,
                'target_value': 0,  # Set to 0 when using custom ranges
                'custom_target_value_low': fit_power_low,
                'custom_target_value_high': fit_power_high
            }
        else:
            # Use heart rate zones
            avg_power = (power_low + power_high) / 2
            hr_zone = self._power_to_heart_rate_zone(avg_power)
            
            return {
                'wkt_step_name': step_name,
                'intensity': Intensity.COOLDOWN,
                'duration_type': duration_type,
                'duration_value': duration_value,
                'target_type': WorkoutStepTarget.HEART_RATE,
                'target_value': hr_zone
            }

    def _parse_intervals(self, element, sport='bike'):
        """Parse interval steps and expand into individual work/rest steps"""
        repeat = int(element.get('Repeat', 1))
        on_duration = int(element.get('OnDuration', 300))  # Duration in seconds
        off_duration = int(element.get('OffDuration', 120))  # Duration in seconds
        on_power = float(element.get('OnPower', 0.85))
        off_power = float(element.get('OffPower', 0.65))
        
        steps = []
        
        for i in range(repeat):
            # Work interval
            if self._should_use_power(sport):
                # Apply buffer to work interval
                on_power_low_watts, on_power_high_watts = self._apply_power_buffer_watts(on_power)
                
                # Convert to FIT format
                fit_on_power_low = self._convert_power_for_fit(on_power_low_watts)
                fit_on_power_high = self._convert_power_for_fit(on_power_high_watts)
                
                work_step = {
                    'wkt_step_name': f'Interval {i+1} - Work',
                    'intensity': Intensity.ACTIVE,
                    'duration_type': WorkoutStepDuration.TIME,
                    'duration_value': on_duration * 1000,
                    'target_type': WorkoutStepTarget.POWER,
                    'target_value': 0,  # Set to 0 when using custom ranges
                    'custom_target_value_low': fit_on_power_low,
                    'custom_target_value_high': fit_on_power_high
                }
            else:
                on_hr_zone = self._power_to_heart_rate_zone(on_power)
                work_step = {
                    'wkt_step_name': f'Interval {i+1} - Work',
                    'intensity': Intensity.ACTIVE,
                    'duration_type': WorkoutStepDuration.TIME,
                    'duration_value': on_duration * 1000,
                    'target_type': WorkoutStepTarget.HEART_RATE,
                    'target_value': on_hr_zone
                }
            steps.append(work_step)
            
            # Recovery interval (only add if not the last repeat)
            if i < repeat - 1:
                if self._should_use_power(sport):
                    # Apply buffer to recovery interval
                    off_power_low_watts, off_power_high_watts = self._apply_power_buffer_watts(off_power)
                    
                    # Convert to FIT format
                    fit_off_power_low = self._convert_power_for_fit(off_power_low_watts)
                    fit_off_power_high = self._convert_power_for_fit(off_power_high_watts)
                    
                    recovery_step = {
                        'wkt_step_name': f'Interval {i+1} - Recovery',
                        'intensity': Intensity.REST,
                        'duration_type': WorkoutStepDuration.TIME,
                        'duration_value': off_duration * 1000,
                        'target_type': WorkoutStepTarget.POWER,
                        'target_value': 0,  # Set to 0 when using custom ranges
                        'custom_target_value_low': fit_off_power_low,
                        'custom_target_value_high': fit_off_power_high
                    }
                else:
                    off_hr_zone = self._power_to_heart_rate_zone(off_power)
                    recovery_step = {
                        'wkt_step_name': f'Interval {i+1} - Recovery',
                        'intensity': Intensity.REST,
                        'duration_type': WorkoutStepDuration.TIME,
                        'duration_value': off_duration * 1000,
                        'target_type': WorkoutStepTarget.HEART_RATE,
                        'target_value': off_hr_zone
                    }
                steps.append(recovery_step)
        
        return steps

    def _parse_steady_state(self, element, sport='bike'):
        """Parse steady state step"""
        duration = int(element.get('Duration', 1200))  # Duration in seconds
        power = float(element.get('Power', 0.75))
        
        if self._should_use_power(sport):
            # Apply buffer to steady state power
            power_low_watts, power_high_watts = self._apply_power_buffer_watts(power)
            
            # Convert to FIT format
            fit_power_low = self._convert_power_for_fit(power_low_watts)
            fit_power_high = self._convert_power_for_fit(power_high_watts)
            
            return {
                'wkt_step_name': 'Steady state',
                'intensity': Intensity.ACTIVE,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,
                'target_type': WorkoutStepTarget.POWER,
                'target_value': 0,  # Set to 0 when using custom ranges
                'custom_target_value_low': fit_power_low,
                'custom_target_value_high': fit_power_high
            }
        else:
            # Use heart rate zones
            hr_zone = self._power_to_heart_rate_zone(power)
            
            return {
                'wkt_step_name': 'Steady state',
                'intensity': Intensity.ACTIVE,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,
                'target_type': WorkoutStepTarget.HEART_RATE,
                'target_value': hr_zone
            }

    def _power_to_heart_rate_zone(self, power_percentage):
        """Convert power percentage to heart rate zone for running"""
        power_pct = power_percentage * 100  # Convert to percentage (0.5 → 50)
        
        print(f"    DEBUG: Converting power {power_percentage} ({power_pct}%) to HR zone")
        
        if power_pct <= 55:  # Easy/Recovery effort → Z1
            zone = 1
        elif power_pct <= 70:  # Aerobic base effort → Z2  
            zone = 2
        elif power_pct <= 85:  # Tempo effort → Z3
            zone = 3
        elif power_pct <= 95:  # Threshold effort → Z4
            zone = 4
        else:  # VO2max+ effort → Z5
            zone = 5
        
        print(f"    DEBUG: Power {power_pct}% → Zone {zone}")
        return zone

    def create_fit_workout(self, workout_data, output_path):
        """Create FIT file from workout data"""
        # Create file ID message
        file_id_message = FileIdMessage()
        file_id_message.type = FileType.WORKOUT
        file_id_message.manufacturer = Manufacturer.GARMIN
        file_id_message.product = 0
        file_id_message.time_created = round(datetime.datetime.now().timestamp() * 1000)
        file_id_message.serial_number = 0x12345678

        # Create workout steps - ensure every step has wkt_step_name
        workout_steps = []
        for i, step_data in enumerate(workout_data['steps']):
            step = WorkoutStepMessage()
            
            # REQUIRED: Set message index
            step.message_index = i
            
            # REQUIRED: Set step name - ensure it's always present
            step.wkt_step_name = step_data.get('wkt_step_name', f'Step {i+1}')
            
            # REQUIRED: Set duration type
            step.duration_type = step_data['duration_type']
            
            # REQUIRED: Set duration value
            step.duration_value = step_data['duration_value']
            
            # REQUIRED: Set target type
            step.target_type = step_data['target_type']
            
            # REQUIRED: Set target value (0 for custom ranges, actual value for single targets)
            step.target_value = step_data['target_value']
            
            # OPTIONAL: Set custom target range
            if 'custom_target_value_low' in step_data:
                step.custom_target_value_low = step_data['custom_target_value_low']
            
            if 'custom_target_value_high' in step_data:
                step.custom_target_value_high = step_data['custom_target_value_high']
            
            # OPTIONAL: Set intensity
            if 'intensity' in step_data:
                step.intensity = step_data['intensity']
            
            # OPTIONAL: Set notes
            if 'notes' in step_data:
                step.notes = step_data['notes']
            
            # OPTIONAL: Set equipment
            if 'equipment' in step_data:
                step.equipment = step_data['equipment']
            
            workout_steps.append(step)

        # Create workout message with wkt_name
        workout_message = WorkoutMessage()
        workout_message.wkt_name = workout_data['name']  # Add workout name
        workout_message.sport = self.sport_mapping.get(workout_data['sport'], Sport.GENERIC)
        workout_message.capabilities = WorkoutCapabilities.TCX  # Add TCX capability
        workout_message.num_valid_steps = len(workout_steps)

        # Build FIT file (removed file_creator_message)
        builder = FitFileBuilder(auto_define=True, min_string_size=50)
        builder.add(file_id_message)
        builder.add(workout_message)
        builder.add_all(workout_steps)

        fit_file = builder.build()
        fit_file.to_file(output_path)
        
        print(f"FIT file created: {output_path}")
        print(f"Workout: {workout_data['name']}")
        print(f"Sport: {workout_data['sport']}")
        print(f"Total steps: {len(workout_steps)}")
        print(f"Power format: {'Absolute watts' if self.use_absolute_power else 'FTP percentage'}")
        print(f"Warmup manual advance: {'Enabled' if self.warmup_manual_advance else 'Disabled'}")
        print(f"Cooldown manual advance: {'Enabled' if self.cooldown_manual_advance else 'Disabled'}")
        if self.force_warmup_power is not None:
            print(f"Forced warmup power: {self.force_warmup_power} ({self.force_warmup_power*100:.0f}%)")
        
        # Print detailed step information
        for i, step_data in enumerate(workout_data['steps'], 1):
            if step_data['duration_type'] == WorkoutStepDuration.OPEN:
                duration_text = "Manual LAP"
            else:
                duration_seconds = step_data['duration_value'] / 1000
                duration_minutes = duration_seconds / 60
                duration_text = f"{duration_minutes:.1f}min"
            
            if step_data['target_type'] == WorkoutStepTarget.POWER:
                if 'custom_target_value_low' in step_data and 'custom_target_value_high' in step_data:
                    # Convert back from FIT format for display
                    if self.use_absolute_power:
                        low_watts = step_data['custom_target_value_low'] - 1000
                        high_watts = step_data['custom_target_value_high'] - 1000
                    else:
                        low_watts = (step_data['custom_target_value_low'] / 10) * self.ftp_watts / 100
                        high_watts = (step_data['custom_target_value_high'] / 10) * self.ftp_watts / 100
                    
                    low_pct = (low_watts / self.ftp_watts) * 100
                    high_pct = (high_watts / self.ftp_watts) * 100
                    print(f"  Step {i}: {step_data['wkt_step_name']} - {duration_text} - {low_watts:.0f}-{high_watts:.0f}W ({low_pct:.0f}%-{high_pct:.0f}% FTP)")
                else:
                    watts = step_data['target_value']
                    if self.use_absolute_power and watts > 1000:
                        watts -= 1000
                    elif not self.use_absolute_power:
                        watts = (watts / 10) * self.ftp_watts / 100
                    
                    pct = (watts / self.ftp_watts) * 100
                    print(f"  Step {i}: {step_data['wkt_step_name']} - {duration_text} - {watts:.0f}W ({pct:.0f}% FTP)")
            else:
                print(f"  Step {i}: {step_data['wkt_step_name']} - {duration_text} - HR Zone {step_data['target_value']}")

    def convert_zwo_to_fit(self, zwo_file_path, output_dir='./'):
        """Convert single zwo file to FIT file"""
        try:
            workout = self.parse_zwo_file(zwo_file_path)
            
            # Create output filename
            safe_name = "".join(c for c in workout['name'] if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_name = safe_name.replace(' ', '_')
            output_filename = f"{safe_name}.fit"
            output_path = os.path.join(output_dir, output_filename)
            
            # Create FIT file
            self.create_fit_workout(workout, output_path)
                
        except Exception as e:
            print(f"Error converting {zwo_file_path} to FIT: {e}")
            raise

    def convert_folder(self, zwo_folder_path, fit_folder_path):
        """Convert all ZWO files in a folder to FIT files"""
        # Ensure the output directory exists
        os.makedirs(fit_folder_path, exist_ok=True)
        
        # Find all .zwo files in the source folder
        zwo_pattern = os.path.join(zwo_folder_path, "*.zwo")
        zwo_files = glob.glob(zwo_pattern)
        
        if not zwo_files:
            print(f"No .zwo files found in {zwo_folder_path}")
            return
        
        print(f"Found {len(zwo_files)} ZWO files to convert:")
        for zwo_file in zwo_files:
            print(f"  - {os.path.basename(zwo_file)}")
        
        print("\n" + "="*60)
        
        # Convert each file
        successful_conversions = 0
        failed_conversions = 0
        
        for zwo_file in zwo_files:
            try:
                print(f"\nConverting: {os.path.basename(zwo_file)}")
                self.convert_zwo_to_fit(zwo_file, fit_folder_path)
                successful_conversions += 1
                print("-" * 40)
                
            except Exception as e:
                print(f"Failed to convert {os.path.basename(zwo_file)}: {e}")
                failed_conversions += 1
        
        # Summary
        print("\n" + "="*60)
        print("CONVERSION SUMMARY:")
        print(f"Total files processed: {len(zwo_files)}")
        print(f"Successful conversions: {successful_conversions}")
        print(f"Failed conversions: {failed_conversions}")
        print(f"Output directory: {fit_folder_path}")


def main():
    # Initialize converter with your FTP in watts and 5% buffer
    converter = zwoToFitConverter(
        ftp_watts=240,  # Your actual FTP
        use_power_for_cycling=True,
        power_buffer_percent=5,  # ±5% buffer on all power targets
        use_absolute_power=True,  # Set to True for absolute watts, False for FTP percentages
        warmup_manual_advance=True,  # Warmup steps wait for LAP button press
        cooldown_manual_advance=False,  # Cooldown steps use timed duration (change to True if desired)
        force_warmup_power=0.5  # Force all warmups to 50% effort (Z1 recovery)
    )
    
    # Define source and destination folders
    zwo_folder = './zwo'  # Folder containing .zwo files
    fit_folder = './fit'  # Folder where .fit files will be saved
    
    # Convert all ZWO files in the folder
    converter.convert_folder(zwo_folder, fit_folder)


if __name__ == "__main__":
    main()