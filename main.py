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
    def __init__(self):
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
            steps = self._parse_workout_steps(workout_element)
        
        return {
            'name': name,
            'description': description,
            'sport': sport_type.lower(),
            'steps': steps
        }

    def _parse_workout_steps(self, workout_element):
        """Parse workout steps and expand intervals"""
        steps = []
        
        for element in workout_element:
            if element.tag == 'Warmup':
                step = self._parse_warmup(element)
                steps.append(step)
                
            elif element.tag == 'Cooldown':
                step = self._parse_cooldown(element)
                steps.append(step)
                
            elif element.tag == 'IntervalsT':
                interval_steps = self._parse_intervals(element)
                steps.extend(interval_steps)
                
            # Add other step types as needed
            elif element.tag == 'SteadyState':
                step = self._parse_steady_state(element)
                steps.append(step)
        
        return steps

    def _parse_warmup(self, element):
        """Parse warmup step"""
        duration = int(element.get('Duration', 600))  # Duration in seconds
        power_low = float(element.get('PowerLow', 0.60))
        power_high = float(element.get('PowerHigh', 0.70))
        
        # Use average power for heart rate zone calculation
        avg_power = (power_low + power_high) / 2
        hr_zone = self._power_to_heart_rate_zone(avg_power)
        
        return {
            'name': 'Warmup',
            'intensity': Intensity.WARMUP,
            'duration_type': WorkoutStepDuration.TIME,
            'duration_value': duration * 1000,  # Convert to milliseconds for FIT
            'target_type': WorkoutStepTarget.HEART_RATE,
            'target_value': hr_zone
        }

    def _parse_cooldown(self, element):
        """Parse cooldown step"""
        duration = int(element.get('Duration', 600))  # Duration in seconds
        power_low = float(element.get('PowerLow', 0.60))
        power_high = float(element.get('PowerHigh', 0.65))
        
        # Use average power for heart rate zone calculation
        avg_power = (power_low + power_high) / 2
        hr_zone = self._power_to_heart_rate_zone(avg_power)
        
        return {
            'name': 'Cooldown',
            'intensity': Intensity.COOLDOWN,
            'duration_type': WorkoutStepDuration.TIME,
            'duration_value': duration * 1000,  # Convert to milliseconds for FIT
            'target_type': WorkoutStepTarget.HEART_RATE,
            'target_value': hr_zone
        }

    def _parse_intervals(self, element):
        """Parse interval steps and expand into individual work/rest steps"""
        repeat = int(element.get('Repeat', 1))
        on_duration = int(element.get('OnDuration', 300))  # Duration in seconds
        off_duration = int(element.get('OffDuration', 120))  # Duration in seconds
        on_power = float(element.get('OnPower', 0.85))
        off_power = float(element.get('OffPower', 0.65))
        
        on_hr_zone = self._power_to_heart_rate_zone(on_power)
        off_hr_zone = self._power_to_heart_rate_zone(off_power)
        
        steps = []
        
        for i in range(repeat):
            # Work interval
            work_step = {
                'name': f'Interval {i+1} - Work',
                'intensity': Intensity.ACTIVE,
                'duration_type': WorkoutStepDuration.TIME,
                'duration_value': on_duration * 1000,  # Convert to milliseconds for FIT
                'target_type': WorkoutStepTarget.HEART_RATE,
                'target_value': on_hr_zone
            }
            steps.append(work_step)
            
            # Recovery interval (only add if not the last repeat)
            if i < repeat - 1:
                recovery_step = {
                    'name': f'Interval {i+1} - Recovery',
                    'intensity': Intensity.REST,
                    'duration_type': WorkoutStepDuration.TIME,
                    'duration_value': off_duration * 1000,  # Convert to milliseconds for FIT
                    'target_type': WorkoutStepTarget.HEART_RATE,
                    'target_value': off_hr_zone
                }
                steps.append(recovery_step)
        
        return steps

    def _parse_steady_state(self, element):
        """Parse steady state step"""
        duration = int(element.get('Duration', 1200))  # Duration in seconds
        power = float(element.get('Power', 0.75))
        
        hr_zone = self._power_to_heart_rate_zone(power)
        
        return {
            'name': 'Steady State',
            'intensity': Intensity.ACTIVE,
            'duration_type': WorkoutStepDuration.TIME,
            'duration_value': duration * 1000,  # Convert to milliseconds for FIT
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

        # Create workout steps
        workout_steps = []
        for i, step_data in enumerate(workout_data['steps']):
            step = WorkoutStepMessage()
            
            # Set message index (REQUIRED COLUMN)
            step.message_index = i
            
            # Set step name if available
            for attr_name in ['wkt_step_name', 'step_name', 'name']:
                if hasattr(step, attr_name):
                    setattr(step, attr_name, step_data['name'])
                    break
            
            # Set intensity
            step.intensity = step_data['intensity']
            
            # Set duration type (REQUIRED)
            step.duration_type = step_data['duration_type']
            
            # Set duration value (REQUIRED)
            for attr_name in ['duration_value', 'durationValue', 'duration']:
                if hasattr(step, attr_name):
                    setattr(step, attr_name, step_data['duration_value'])
                    break
            
            # Set target type and value
            step.target_type = step_data['target_type']
            step.target_value = step_data['target_value']
            
            # Set secondary target value (REQUIRED COLUMN - usually 0)
            step.secondary_target_value = 0
            
            # Set weight display unit (REQUIRED COLUMN)
            step.weight_display_unit = 0  # 0 = kilogram
            
            workout_steps.append(step)

        # Create workout message with all required fields
        workout_message = WorkoutMessage()
        
        # Try different possible attribute names for workout name
        workout_name = workout_data['name']
        for attr_name in ['wkt_name', 'workout_name', 'name', 'wktName']:
            if hasattr(workout_message, attr_name):
                setattr(workout_message, attr_name, workout_name)
                break
        
        # Try multiple possible field names for message_index
        for attr_name in ['message_index', 'messageIndex', 'msg_index']:
            if hasattr(workout_message, attr_name):
                setattr(workout_message, attr_name, 0)
                break
        
        # Set other required fields
        workout_message.sport = self.sport_mapping.get(workout_data['sport'], Sport.GENERIC)
        workout_message.num_valid_steps = len(workout_steps)
        
        # Try to set sub_sport
        for attr_name in ['sub_sport', 'subSport']:
            if hasattr(workout_message, attr_name):
                setattr(workout_message, attr_name, 0)  # 0 = generic
                break
        
        # Try to set capabilities
        for attr_name in ['capabilities']:
            if hasattr(workout_message, attr_name):
                setattr(workout_message, attr_name, 32)  # TCX capability
                break

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
        print(f"Total steps: {len(workout_steps)}")
        
        # Print detailed step information
        for i, step_data in enumerate(workout_data['steps'], 1):
            duration_seconds = step_data['duration_value'] / 1000
            print(f"  Step {i}: {step_data['name']} - {duration_seconds:.0f}s - HR Zone {step_data['target_value']}")

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
    converter = zwoToFitConverter()
    
    # Define source and destination folders
    zwo_folder = './zwo'  # Folder containing .zwo files
    fit_folder = './fit'  # Folder where .fit files will be saved
    
    # Convert all ZWO files in the folder
    converter.convert_folder(zwo_folder, fit_folder)


if __name__ == "__main__":
    main()