class Lesson:

    def __init__(self, text):
        """Initialize a lesson object with all needed variables to add onto the Excel"""
        self.text = text
        self.info = {}
        # Text fixes for course_name
        self._text_fixes()
        # Fill info
        key_names = ['course_name', 'professor', 'area_name', 'startTime', 'endTime']
        self._extract_variables(key_names)
        self._extract_variables(['daysOfWeek'], ("[", "]"))
        # Find how many hours the lesson lasts
        self._get_time()

    def _extract_variables(self, key_names, value_mark=("'", "'")):
        """Add new info that exists in text"""
        # Info structure key_name: 'value', value_mark tuple sets the start and end of a value
        # Extracting in a loop
        for key_name in key_names:
            # Find where the key starts and save text starting there
            key_start = self.text.find(key_name)
            value = self.text[key_start:]
            # Find where the value of the key starts
            value_start = value.find(value_mark[0])
            value = value[value_start + 1:]
            # Find where the value of the key ends
            value_end = value.find(value_mark[1])
            value = value[:value_end]
            # Save key:value pair
            if len(value) > 0:
                if key_name == 'daysOfWeek':
                    self.info[key_name] = int(value)
                elif key_name == 'course_name':
                    for fix in self.fixes:
                        if value.find(fix) != -1:
                            value = value.replace(fix, 'KAI')
                    self.info[key_name] = value.upper()
                else:
                    self.info[key_name] = value
            else:
                self.info[key_name] = ' '

    def _text_fixes(self):
        """All text fixes needed for a course_name"""
        self.fixes = [
            '&amp;amp;',
            '&amp;',
            'amp;',
        ]

    def _get_time(self):
        """Get how many hours the lesson lasts"""
        # Extract hour int value
        hour_start = int(self.info['startTime'][:2])
        hour_end = int(self.info['endTime'][:2])
        # Save lasting hours
        self.info['lasting'] = hour_end - hour_start

    def print_info(self):
        """Print lesson's saved info"""
        print('Printing info:')
        for key, value in self.info.items():
            print(f"\tkey: {key}")
            print(f"\tvalue: {value}")


def get_course_name(lesson_inst):
    """Gets the name of a lesson"""
    return lesson_inst.info['course_name']
