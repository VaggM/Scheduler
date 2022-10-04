import PySimpleGUI as sg
import os
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

        # Data for translating file names
        self._get_translations()

        # Add blank line
        self.new_text('')
        self.add_line()

        # Get all available files within urls folder
        self.new_text('Διαθέσιμα προγράμματα: \t')
        self._get_available_urls()
        self.new_combo_list(self.named_urls, key='urls')
        # self.new_text('   ')
        # self.new_button('Ενημέρωση\nπρογραμμάτων', key='update', size=(12, 2))
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

        self.current_folder = os.getcwd() + '\\schedules'
        self.new_text('Φάκελος προορισμού:   ')
        self.new_textfield(text=self.current_folder, key='folder', disabled=False)
        self.new_browser_folder()
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

    def _get_translations(self):
        """Translate codes into department names"""
        self.translations = [
            # ΣΧΟΛΗ ΔΗΜΟΣΙΑΣ ΥΓΕΙΑΣ
            ['pch',         'Δημόσιας και Κοινοτικής Υγείας'],
            ['php',         'Πολιτικών Δημόσιας Υγείας'],

            # ΣΧΟΛΗ ΔΙΟΙΚΗΤΙΚΩΝ, ΟΙΚΟΝΟΜΙΚΩΝ & ΚΟΙΝΩΝΙΚΩΝ ΕΠΙΣΤΗΜΩΝ
            ['ecec',        'Αγωγής και Φροντίδας στην Πρώιμη Παιδική Ηλικία'],
            ['alis',        'Αρχειονομίας, Βιβλιοθηκονομίας και Συστημάτων Πληροφόρησης'],
            ['ba',          'Διοίκησης Επιχειρήσεων'],
            ['tourism',     'Διοίκησης Τουρισμού'],
            ['sw',          'Κοινωνικής Εργασίας'],
            ['accfin',      'Λογιστικής και Χρηματοοικονομικής'],

            # ΣΧΟΛΗ ΕΠΙΣΤΗΜΩΝ ΤΡΟΦΙΜΩΝ
            ['fst',         'Επιστήμης και Τεχνολογίας Τροφίμων'],
            ['wvbs',        'Επιστημών Οίνου, Αμπέλου και Ποτών'],

            # ΣΧΟΛΗ ΕΠΙΣΤΗΜΩΝ ΥΓΕΙΑΣ & ΠΡΟΝΟΙΑΣ
            ['bisc',        'Βιοϊατρικών Επιστημών'],
            ['ot',          'Εργοθεραπείας'],
            ['midw',        'Μαιευτικής'],
            ['nurs',        'Νοσηλευτικής'],
            ['phys',        'Φυσικοθεραπείας'],

            # ΣΧΟΛΗ ΕΦΑΡΜΟΣΜΕΝΩΝ ΤΕΧΝΩΝ & ΠΟΛΙΤΙΣΜΟΥ
            ['gd',          'Γραφιστικής και Οπτικής Επικοινωνίας'],
            ['ia',          'Εσωτερικής Αρχιτεκτονικής'],
            ['cons',        'Συντήρησης Αρχαιοτήτων και Έργων Τέχνης'],
            ['phaa',        'Φωτογραφίας και Οπτικοακουστικών Τεχνών'],

            # ΣΧΟΛΗ ΜΗΧΑΝΙΚΩΝ
            ['eee',         'Ηλεκτρολόγων και Ηλεκτρονικών Μηχανικών'],
            ['bme',         'Μηχανικών Βιοϊατρικής'],
            ['idpe',        'Μηχανικών Βιομηχανικής Σχεδίασης και Παραγωγής'],
            ['ice',         'Μηχανικών Πληροφορικής και Υπολογιστών'],
            ['geo',         'Μηχανικών Τοπογραφίας και Γεωπληροφορικής'],
            ['mech',        'Μηχανολόγων Μηχανικών'],
            ['na',          'Ναυπηγών Μηχανικών'],
            ['civ',         'Πολιτικών Μηχανικών'],
        ]

    def _get_available_urls(self):
        """Get all file directories within urls folder"""
        self.available_urls = []
        self.named_urls = []
        directory = os.getcwd() + "\\urls"
        try:
            os.mkdir(directory)
        except FileExistsError:
            pass
        files = os.listdir(directory)
        for file in files:
            index = self._formed_name(file)
            self.available_urls.insert(index, directory + "\\" + file)

    def _formed_name(self, name):
        """Create a formed name to display to the user for each url file"""
        # from "eee.spring.2022-2023"
        # to "Εαρινό εξάμηνο 2022-2023, Τμήμα Ηλεκτρολόγων και Ηλεκτρονικών Μηχανικών
        translates = self.translations

        part1 = name.find('.')
        part1 = name[:part1]
        for translate in translates:
            if part1 == translate[0]:
                part1 = translate[1]
                break

        name = name[name.find('.')+1:]
        part2 = name.find('.')
        part2 = name[:part2]
        if part2 == 'winter':
            part2 = 'Χειμερινό'
        elif part2 == 'spring':
            part2 = 'Εαρινό'

        name = name[name.find('.')+1:]
        part3 = name[:name.find('.')]

        output = f" {part2} εξάμηνο {part3}, Τμήμα {part1}"
        self.named_urls.append(output)
        self.named_urls.sort()
        # return the index of sorted output
        return self.named_urls.index(output)

    def main(self):
        """Main method handling window events"""
        # Events correspond to keys
        while True:
            event, values = self.window.read()
            if event == sg.WIN_CLOSED or event == "Exit":
                break
            elif event == 'folder':
                self.current_folder = values['folder']
                self.text_update('folder', values['folder'])
            elif event == 'complete':
                try:
                    if self.current_folder != (os.getcwd() + '\\schedules'):
                        os.listdir(self.current_folder)
                    self._complete()
                except FileNotFoundError:
                    sg.Popup('Δεν υπάρχει αυτός ο φάκελος!', font=(50, 15), title='Error')
            elif event == 'urls':
                try:
                    formed_name = values['urls']
                    index = self.named_urls.index(formed_name)
                    filename = self.available_urls[index]
                    self._explore_file(filename)
                    self._list1_fill()
                    self._list2_empty()
                except ValueError:
                    sg.Popup('Υπάρχει λάθος με το συγκεκριμένο αρχείο!', font=(50, 15), title='Error')
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

    def _list2_empty(self):
        """Empty list2"""
        self.names = []
        self.window['list2'].Update(values=self.names)
        self.text_update('text2', f"\t\tΕπιλεγμένα μαθήματα: {len(self.names)}")

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
            except xlsxwriter.exceptions.FileCreateError:
                sg.Popup('Κλείσε το παλιό Excel πρώτα!', font=(50, 15), title='Error')
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
        create_excel_schedule(desired_lessons, self.current_folder)
