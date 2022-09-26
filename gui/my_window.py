import PySimpleGUI as sg
import xlsxwriter.exceptions

from gui.gui_window import GuiWindow
from excel_creation.create_excel import create_excel_schedule


class MyWindow(GuiWindow):
    """Using GuiWindow functions to create a usable layout"""

    def __init__(self, lessons, url_lessons):
        """Initialize layout and create window"""
        # Add title
        super().__init__('Δημιουργία Excel')

        # List1 and List2 data
        self.lessons = lessons
        self.names = []

        # Html data
        self.url_lessons = url_lessons

        # Window text
        self.new_text(f"\t\tΔιαθέσιμα μαθήματα: {len(self.lessons)}")
        self.new_text("\t\t\t")
        self.new_text(f"\t\tΕπιλεγμένα μαθήματα: {len(self.names)}", key='choice_nums')
        self.add_line()

        # Create lists next to each other
        self.new_list(self.lessons, 'list1')
        self.new_list(self.names, 'list2')
        self.add_line()

        # Add blank line
        self.new_text('')
        self.add_line()

        # Add excel creation button
        self.new_text("\t\t\t\t\t")
        self.new_button('Δημιουργία Excel', key='complete')
        self.add_line()

        # Create window
        self.create_window()

    def main(self):
        """Main method handling window events"""

        # Events correspond to keys
        while True:
            event, values = self.window.read()
            if event == sg.WIN_CLOSED or event == "Exit":
                break
            elif event == 'complete':
                completion = self._complete()
                if completion:
                    break
            elif event[:5] == 'radio':
                self.swap_radio_state(event)
            elif event == 'list1':
                self.list2_add(values['list1'])
            elif event == 'list2':
                self.list2_remove(values['list2'])

        self.window.close()

    def list2_add(self, names):
        """Add names to list2 from list1 and update window"""

        for name in names:
            if name not in self.names:
                self.names.append(name)
                self.window['list2'].Update(values=self.names)

                self.text_update('choice_nums', f"\t\tΕπιλεγμένα μαθήματα: {len(self.names)}")

    def list2_remove(self, names):
        """Remove names from list2 and update the window"""

        for name in names:
            self.names.remove(name)
            self.window['list2'].Update(values=self.names)

            self.text_update('choice_nums', f"\t\tΕπιλεγμένα μαθήματα: {len(self.names)}")

    def _complete(self):
        """Create excel from list2 names"""
        if len(self.names) > 0:
            try:
                self._excel_creation()
                sg.Popup('Excel successfully created!', font=(50, 15), title='Complete')
                return True
            except xlsxwriter.exceptions.FileCreateError:
                sg.Popup('Please close your old Excel file!', font=(50, 15), title='Error')
                return False
        else:
            sg.Popup('Please choose a lesson first!', font=(50, 15), title='Error')

    def _excel_creation(self):
        """Prepare desired_lessons variable to send to the Excel creation function"""
        # Get all lesson names
        lesson_names = []
        for les in self.url_lessons:
            lesson_names.append(les.info['course_name'])
        # Add info on desired lessons
        desired_lessons = []
        for index, name in enumerate(lesson_names):
            if name in self.names:
                info = [
                    self.url_lessons[index].info['course_name'],
                    self.url_lessons[index].info['professor'],
                    self.url_lessons[index].info['area_name'],
                    self.url_lessons[index].info['daysOfWeek'],
                    self.url_lessons[index].info['startTime'],
                    self.url_lessons[index].info['lasting'],
                ]
                desired_lessons.append(info)
        # Send info to Excel creation
        create_excel_schedule(desired_lessons)
