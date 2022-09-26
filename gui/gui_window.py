import PySimpleGUI as sg
import os


class GuiWindow:
    """Class to make PySimpleGui commands easier to use"""

    def __init__(self, title):
        """Initialize a gui window"""
        # Picks a window theme based on all PySimpleGUI themes
        sg.theme('DarkGray12')
        # Initialize variables
        self.title = title
        self.layout = []
        self.stack = []

    def new_text(self, text, key=None):
        """Add new text element to line stack"""
        self.stack.append(sg.T(text, key=key))

    def new_button(self, text, key='default', size=(35, 1), visible=True):
        """Add new button element to line stack"""
        self.stack.append(sg.Button(text, size=size, key=key, visible=visible))

    def new_radio(self, text, key):
        """Add new radio button element to line stack"""
        key = str(key)
        self.stack.append(sg.Radio(text, key, default=False, key=key, enable_events=True))

    def new_list(self, names, key):
        """Add new listbox element to line stack"""
        self.stack.append(sg.Listbox(values=names, size=(60, 20), key=key, enable_events=True))

    def new_browser(self):
        """Add new browse button element to line stack"""
        working_directory = os.getcwd()
        self.stack.append(sg.FileBrowse(button_text='Αναζήτηση', initial_folder=working_directory))

    def new_textfield(self, key, disabled=False):
        """Add new textfield element to line stack"""
        self.stack.append(sg.InputText(key=key, enable_events=True, disabled=disabled))

    def add_line(self):
        """Add a new line with all elements within the stack and reset stack"""
        self.layout.append(self.stack)
        self.stack = []

    def hide_row(self, elem_key):
        """Hides a row of an element"""
        self.window[elem_key].hide_row()

    def reveal_row(self, elem_key):
        """Reveals a row of an element"""
        self.window[elem_key].unhide_row()

    def visibility(self, elem_key, visible=True):
        """Set element visible/invisible"""
        self.window[elem_key].Update(visible=visible)

    def radio_text_change(self, elem_key, text):
        """Change text of a radio button"""
        self.window[elem_key].Update(text=text)

    def text_update(self, elem_key, text):
        """Update a text element"""
        self.window[elem_key].update(text)

    def swap_radio_state(self, elem_key):
        """Swap state of a radio button"""
        value = self.window[elem_key].Value
        if not value:
            self.set_radio_state(elem_key, True)
        else:
            self.set_radio_state(elem_key, False)

    def set_radio_state(self, elem_key, state):
        """Set state of a radio button (True/False)"""
        self.window[elem_key].Value = state
        self.window[elem_key].Update(value=state)

    def get_radio_value(self, elem_key):
        """Get radio button (True/False) value"""
        return self.window[elem_key].Value

    def get_radio_text(self, elem_key):
        """Get radio button text"""
        return self.window[elem_key].Text

    def create_window(self):
        """Create the main window"""
        self.window = sg.Window(self.title, self.layout, size=(900, 550))
