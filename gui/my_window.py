import PySimpleGUI as sg
import xlsxwriter.exceptions

from gui.gui_window import GuiWindow
from html_info.html_explore import get_html_lessons
from html_info.lesson import get_course_name
from excel_creation.create_excel import create_excel_schedule


class MyWindow(GuiWindow):
    """Using GuiWindow functions to create a usable layout"""

    def __init__(self):
        """Initialize layout and create window"""
        # Add title
        super().__init__('Επιλογή μαθημάτων')

        # List1 and List2 data
        self.lessons = []
        self.names = []

        # Get filename and its data
        self.new_text('Επιλογή αρχείου: ')
        self.new_textfield('filename', disabled=True)
        self.new_browser()
        self.add_line()

        # Add blank line
        self.new_text('')
        self.add_line()

        # Window text
        self.new_text(f"\t\tΔιαθέσιμα μαθήματα: {len(self.lessons)}", key='text1')
        self.new_text("\t\t\t")
        self.new_text(f"\t\tΕπιλεγμένα μαθήματα: {len(self.names)}", key='text2')
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
                # if completion:
                #    break
            elif event == 'filename':
                try:
                    filename = values['filename']
                    self._explore_file(filename)
                    self._list1_fill()
                except ValueError:
                    sg.Popup('Λάθος αρχείο!', font=(50, 15), title='Error')
            elif event == 'list1':
                self._list2_add(values['list1'])
            elif event == 'list2':
                self._list2_remove(values['list2'])

        self.window.close()

    def _explore_file(self, filename):
        """Get all lessons from the selected file"""
        self.url_lessons = get_html_lessons(filename)
        self.url_lessons.sort(key=get_course_name)

    def _list1_fill(self):
        """Find all unique lesson names and update list1, text1"""
        # Find all unique names
        self.lessons = []
        for lesson in self.url_lessons:
            name = get_course_name(lesson)
            if name not in self.lessons:
                self.lessons.append(name)
        # Update list1, text1
        self.window['list1'].Update(values=self.lessons)
        self.text_update('text1', f"\t\tΔιαθέσιμα μαθήματα: {len(self.lessons)}")

    def _list2_add(self, names):
        """Add names to list2 from list1 and update window"""
        for name in names:
            if name not in self.names:
                self.names.append(name)
                self.window['list2'].Update(values=self.names)

                self.text_update('text2', f"\t\tΕπιλεγμένα μαθήματα: {len(self.names)}")

    def _list2_remove(self, names):
        """Remove names from list2 and update the window"""
        for name in names:
            self.names.remove(name)
            self.window['list2'].Update(values=self.names)

            self.text_update('text2', f"\t\tΕπιλεγμένα μαθήματα: {len(self.names)}")

    def _complete(self):
        """Create excel from list2 names"""
        if len(self.names) > 0 and len(self.lessons) > 0:
            try:
                self._excel_creation()
                sg.Popup('Το Excel δημιουργήθηκε επιτυχώς!', font=(50, 15), title='Complete')
                return True
            except xlsxwriter.exceptions.FileCreateError:
                sg.Popup('Κλείσε το παλιό Excel πρώτα!', font=(50, 15), title='Error')
                return False
        elif len(self.names) == 0:
            sg.Popup('Διάλεξε ένα μάθημα πρώτα!', font=(50, 15), title='Error')
        elif len(self.lessons) == 0:
            sg.Popup('Διάλεξε ένα αρχείο πρώτα!', font=(50, 15), title='Error')

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
