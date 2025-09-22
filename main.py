import datetime
import xml.etree.ElementTree as ET
import os
import glob
from fit_tool.fit_file_builder import FitFileBuilder
from fit_tool.profile.messages.file_id_message import FileIdMessage
from fit_tool.profile.messages.workout_message import WorkoutMessage
from fit_tool.profile.messages.workout_step_message import WorkoutStepMessage
from fit_tool.profile.messages.file_creator_message import FileCreatorMessage
from fit_tool.profile.profile_type import Sport, Intensity, WorkoutStepDuration, WorkoutStepTarget, Manufacturer, FileType


class zwoToFitConverter:
    def __init__(self, ftp_watts=250, use_power_for_cycling=True):
        """
        Initialize converter
        
        Args:
            ftp_watts: Your FTP in watts (used to convert percentages to absolute power)
            use_power_for_cycling: If True, use power targets for cycling workouts
        """
        self.ftp_watts = ftp_watts
        self.use_power_for_cycling = use_power_for_cycling
        
        # Mapping zwo sport types to FIT sport types
        self.sport_mapping = {
            'run': Sport.RUNNING,
            'bike': Sport.CYCLING,
            'cycling': Sport.CYCLING,
            'swim': Sport.SWIMMING,
            'other': Sport.GENERIC
        }
        
        # Power zone to heart rate zone mapping (approximate)
        self.power_to_hr_zone = {
            0.50: 1,  # Recovery
            0.60: 2,  # Aerobic base
            0.65: 2,  # Aerobic base
            0.70: 3,  # Aerobic
            0.75: 3,  # Aerobic
            0.80: 4,  # Threshold
            0.85: 4,  # Threshold
            0.90: 4,  # Threshold
            0.95: 4,  # Threshold/VO2max
            1.00: 5,  # VO2max
            1.05: 5,  # VO2max
            1.10: 5,  # VO2max
            1.15: 5,  # Neuromuscular
            1.20: 5   # Neuromuscular
        }

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

    def _parse_warmup(self, element, sport='bike'):
        """Parse warmup step"""
        duration = int(element.get('Duration', 600))  # Duration in seconds
        power_low = float(element.get('PowerLow', 0.60))
        power_high = float(element.get('PowerHigh', 0.70))
        
        if self._should_use_power(sport):
            # Use power targets for cycling
            power_low_watts = int(power_low * self.ftp_watts)
            power_high_watts = int(power_high * self.ftp_watts)
            
            return {
                'wkt_step_name': 'Warmup',
                'intensity': Intensity.WARMUP,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,  # Convert to milliseconds for FIT
                'target_type': WorkoutStepTarget.POWER,
                'target_value': (power_low_watts + power_high_watts) // 2,
                'custom_target_value_low': power_low_watts,
                'custom_target_value_high': power_high_watts
            }
        else:
            # Use heart rate zones (original behavior)
            avg_power = (power_low + power_high) / 2
            hr_zone = self._power_to_heart_rate_zone(avg_power)
            
            return {
                'wkt_step_name': 'Warmup',
                'intensity': Intensity.WARMUP,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,
                'target_type': WorkoutStepTarget.HEART_RATE,
                'target_value': hr_zone
            }

    def _parse_cooldown(self, element, sport='bike'):
        """Parse cooldown step"""
        duration = int(element.get('Duration', 600))  # Duration in seconds
        power_low = float(element.get('PowerLow', 0.60))
        power_high = float(element.get('PowerHigh', 0.65))
        
        if self._should_use_power(sport):
            # Use power targets for cycling
            power_low_watts = int(power_low * self.ftp_watts)
            power_high_watts = int(power_high * self.ftp_watts)
            
            return {
                'wkt_step_name': 'Cooldown',
                'intensity': Intensity.COOLDOWN,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,
                'target_type': WorkoutStepTarget.POWER,
                'target_value': (power_low_watts + power_high_watts) // 2,
                'custom_target_value_low': power_low_watts,
                'custom_target_value_high': power_high_watts
            }
        else:
            # Use heart rate zones
            avg_power = (power_low + power_high) / 2
            hr_zone = self._power_to_heart_rate_zone(avg_power)
            
            return {
                'wkt_step_name': 'Cooldown',
                'intensity': Intensity.COOLDOWN,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,
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
                on_power_watts = int(on_power * self.ftp_watts)
                work_step = {
                    'wkt_step_name': f'Interval {i+1} - Work',
                    'intensity': Intensity.ACTIVE,
                    'duration_type': WorkoutStepDuration.TIME,
                    'duration_value': on_duration * 1000,
                    'target_type': WorkoutStepTarget.POWER,
                    'target_value': on_power_watts,
                    'custom_target_value_low': on_power_watts,
                    'custom_target_value_high': on_power_watts
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
                    off_power_watts = int(off_power * self.ftp_watts)
                    recovery_step = {
                        'wkt_step_name': f'Interval {i+1} - Recovery',
                        'intensity': Intensity.REST,
                        'duration_type': WorkoutStepDuration.TIME,
                        'duration_value': off_duration * 1000,
                        'target_type': WorkoutStepTarget.POWER,
                        'target_value': off_power_watts,
                        'custom_target_value_low': off_power_watts,
                        'custom_target_value_high': off_power_watts
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
            # Use power targets for cycling
            power_watts = int(power * self.ftp_watts)
            
            return {
                'wkt_step_name': f'Steady State ({int(power*100)}% FTP)',
                'intensity': Intensity.ACTIVE,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,
                'target_type': WorkoutStepTarget.POWER,
                'target_value': power_watts,
                'custom_target_value_low': power_watts,
                'custom_target_value_high': power_watts
            }
        else:
            # Use heart rate zones
            hr_zone = self._power_to_heart_rate_zone(power)
            
            return {
                'wkt_step_name': 'Steady State',
                'intensity': Intensity.ACTIVE,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': duration * 1000,
                'target_type': WorkoutStepTarget.HEART_RATE,
                'target_value': hr_zone
            }

    def _power_to_heart_rate_zone(self, power_percentage):
        """Convert power percentage to heart rate zone"""
        # Find the closest power percentage in our mapping
        closest_power = min(self.power_to_hr_zone.keys(), 
                          key=lambda x: abs(x - power_percentage))
        return self.power_to_hr_zone[closest_power]

    def create_fit_workout(self, workout_data, output_path):
        """Create FIT file from workout data"""
        # Create file ID message
        file_id_message = FileIdMessage()
        file_id_message.type = FileType.WORKOUT
        file_id_message.manufacturer = Manufacturer.GARMIN
        file_id_message.product = 0
        file_id_message.time_created = round(datetime.datetime.now().timestamp() * 1000)
        file_id_message.serial_number = 0x12345678

        # Create file creator message
        file_creator_message = FileCreatorMessage()
        file_creator_message.hardware_version = 0
        file_creator_message.software_version = 0

        # Create workout steps using only allowed fields
        workout_steps = []
        for i, step_data in enumerate(workout_data['steps']):
            step = WorkoutStepMessage()
            
            # REQUIRED: Set message index
            step.message_index = i
            
            # OPTIONAL: Set step name
            if 'wkt_step_name' in step_data:
                step.wkt_step_name = step_data['wkt_step_name']
            
            # REQUIRED: Set duration type
            step.duration_type = step_data['duration_type']
            
            # REQUIRED: Set duration value
            step.duration_value = step_data['duration_value']
            
            # REQUIRED: Set target type
            step.target_type = step_data['target_type']
            
            # OPTIONAL: Set target value
            if 'target_value' in step_data:
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

        # Create workout message
        workout_message = WorkoutMessage()
        workout_message.wkt_name = workout_data['name']
        workout_message.sport = self.sport_mapping.get(workout_data['sport'], Sport.GENERIC)
        workout_message.num_valid_steps = len(workout_steps)

        # Build FIT file
        builder = FitFileBuilder(auto_define=True, min_string_size=50)
        builder.add(file_id_message)
        builder.add(file_creator_message)
        builder.add(workout_message)
        builder.add_all(workout_steps)

        fit_file = builder.build()
        fit_file.to_file(output_path)
        
        print(f"FIT file created: {output_path}")
        print(f"Workout: {workout_data['name']}")
        print(f"Sport: {workout_data['sport']}")
        print(f"Total steps: {len(workout_steps)}")
        
        # Print detailed step information
        for i, step_data in enumerate(workout_data['steps'], 1):
            duration_seconds = step_data['duration_value'] / 1000
            duration_minutes = duration_seconds / 60
            
            if step_data['target_type'] == WorkoutStepTarget.POWER:
                if 'custom_target_value_low' in step_data and 'custom_target_value_high' in step_data:
                    print(f"  Step {i}: {step_data['wkt_step_name']} - {duration_minutes:.1f}min - {step_data['custom_target_value_low']}-{step_data['custom_target_value_high']}W")
                else:
                    print(f"  Step {i}: {step_data['wkt_step_name']} - {duration_minutes:.1f}min - {step_data['target_value']}W")
            else:
                print(f"  Step {i}: {step_data['wkt_step_name']} - {duration_minutes:.1f}min - HR Zone {step_data['target_value']}")

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
    # Initialize converter with your FTP in watts
    # Set your actual FTP here!
    converter = zwoToFitConverter(
        ftp_watts=240,  # Replace with your actual FTP
        use_power_for_cycling=True
    )
    
    # Define source and destination folders
    zwo_folder = './zwo'  # Folder containing .zwo files
    fit_folder = './fit'  # Folder where .fit files will be saved
    
    # Convert all ZWO files in the folder
    converter.convert_folder(zwo_folder, fit_folder)


if __name__ == "__main__":
    main()