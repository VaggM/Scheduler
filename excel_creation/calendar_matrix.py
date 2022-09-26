class SpaceTakenError(Exception):
    """There is another subject written"""
    pass


class CalendarMatrix:
    """Class to keep the matrix that will be written into an Excel"""

    def __init__(self, time_start=9, time_end=21, days=('Δευτέρα', 'Τρίτη', 'Τετάρτη', 'Πέμπτη', 'Παρασκευή')):
        """Initialize base matrix and needed variables"""
        # Variables
        self.calendar = []
        self.columns = 1 + len(days)
        self.rows = 1 + time_end - time_start
        self.days = days
        # Variables needed for styling
        self.column_width = 0
        self.row_height = 0

        # Initialize before adding lessons
        self._set_matrix_limits()
        self._set_days_row(days)
        self._set_times_column(time_start)

        # Variables to use later
        self._future_variables()

    def _future_variables(self):
        """Initialize future variables from init"""
        self.merge_cells = []

    def _set_matrix_limits(self):
        """Fill matrix with '' strings"""
        for i in range(self.rows):
            row = []
            for j in range(self.columns):
                row.append('')
            self.calendar.append(row)

    def _set_days_row(self, days):
        """Set strings on first matrix row about days"""
        # First element (0,0)
        self.calendar[0][0] = 'Ώρες \\ Ημέρες'
        # Adding day strings
        for i in range(len(days)):
            self.calendar[0][i + 1] = days[i]

    def _set_times_column(self, time_start):
        """Set strings on first matrix column about times"""
        for i in range(self.rows - 1):
            time = f'{i + time_start}'
            if len(time) == 1:
                time = '0' + time
            self.calendar[i + 1][0] = time

    def write_lesson(self, lesson_text, day, time, lasting=2):
        """Write a lesson within the matrix"""
        x = 0
        y = 0
        day = self.days[day - 1]
        start = time
        # Find day
        for col in self.calendar[0]:
            if col == day:
                x = self.calendar[0].index(col)
                break
        # Find time
        for row in self.calendar:
            if row[0] == start:
                y = self.calendar.index(row)
        # Find empty space to write lesson
        self._check_space_taken(x, y, lesson_text, lasting)

    def _check_space_taken(self, x, y, lesson_text, lasting):
        """Tries to find free space for the lesson and creates a new column if there is none"""
        try:
            for i in range(lasting):
                if self.calendar[y + i][x] != '':
                    raise SpaceTakenError
                # Free space found, writing lesson
            for i in range(lasting):
                self.calendar[y + i][x] = lesson_text
        except SpaceTakenError:
            # Add column if needed
            condition = True
            if x + 1 < len(self.calendar[0]):
                condition = self.calendar[0][x] != self.calendar[0][x + 1]
            if condition:
                for row in self.calendar:
                    if row == self.calendar[0]:
                        day = f"{row[x]}"
                        row.insert(x + 1, day)
                    else:
                        row.insert(x + 1, '')
            # Try again
            self._check_space_taken(x + 1, y, lesson_text, lasting)

    def find_width_height(self):
        """Set width and height"""
        self.column_width = 23
        self.row_height = 40

    def get_merge_cells(self):
        """Find duplicates to merge them after"""
        # Find first row duplicates
        first_row = self.calendar[0]
        j = 0
        while True:
            text = first_row[j]
            k = 0
            if j + 1 < len(first_row):
                while text == first_row[j + k + 1]:
                    k += 1
                    if j + k + 1 >= len(first_row):
                        break
            if k != 0:
                self.merge_cells.append(((0, j), (0, j + k), text))
                j = j + k
            j += 1
            if j >= len(first_row):
                break

        # Find vertical duplicates
        for j in range(1, len(first_row)):
            i = 1
            while True:
                row = self.calendar[i]
                if row != self.calendar[0]:
                    text = row[j]
                    k = 0
                    if i + 1 < len(self.calendar):
                        while text == self.calendar[i + k + 1][j]:
                            k += 1
                            if i + k + 1 >= len(self.calendar):
                                break
                    if k != 0:
                        self.merge_cells.append(((i, j), (i + k, j), text))
                        i = i + k
                i += 1
                if i >= len(self.calendar):
                    break

    def empty_removal(self):
        """Removes all unused rows until you find the first lesson from top to bottom"""
        rows_to_delete = []

        for index, row in enumerate(self.calendar[1:]):
            # Find if there is a lesson in the row
            empty_row = True
            for j in row[1:]:
                if j != '':
                    empty_row = False
                    break
            # Add more rows to delete or break if there is a lesson in current row
            if empty_row:
                rows_to_delete.append(index)
            else:
                break
        # Start deleting rows
        for index, row_num in enumerate(rows_to_delete):
            self.calendar.remove(self.calendar[row_num + 1 - index])

    def empty_removal_reverse(self):
        """Removes all unused rows until you find the first lesson from bottom to top"""
        rows_to_delete = []

        for index in range(len((self.calendar[1:]))):
            # Reversed search
            row = self.calendar[-1 - index]
            # Find if there is a lesson in the row
            empty_row = True
            for j in row[1:]:
                if j != '':
                    empty_row = False
                    break
            # Add more rows to delete or break if there is a lesson in current row
            if empty_row:
                rows_to_delete.append(-1 - index)
            else:
                break
        # Start deleting rows
        for index, row_num in enumerate(rows_to_delete):
            self.calendar.remove(self.calendar[row_num+index])
