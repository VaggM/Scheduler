import xlsxwriter
from excel_creation.calendar_matrix import CalendarMatrix
import os


def create_excel_schedule(lessons):
    """Write all lessons info within an Excel schedule using an edited matrix for reference"""
    # Create a calendar instance
    cldr = CalendarMatrix()
    # Write all lessons info
    for lesson in lessons:
        # Variables to write the lesson
        text = lesson[0] + '\n' + lesson[1] + '\n' + lesson[2]
        day = lesson[3]
        time_start = lesson[4][:2]
        lasting = lesson[5]
        # Write on calendar
        cldr.write_lesson(text, day, time_start, lasting)

    # example: cldr.writeLesson('Κυκλώματα ΙΙ\nΤσεκούρας Γεώργιος\nZB001', 'Παρασκευή', '15', 3)

    # Fix the saved calendar
    cldr.check_saturday()
    cldr.empty_removal()
    cldr.empty_removal_reverse()
    cldr.find_width_height()
    cldr.get_merge_cells()

    working_dir = os.getcwd() + '\\schedules'
    try:
        os.mkdir(working_dir)
    except FileExistsError:
        pass
    file_names = os.listdir(working_dir)
    i = 0
    while True:
        filename = f"Schedule({i}).xlsx"
        if filename in file_names:
            i += 1
        else:
            break

    # Excel edit
    directory = working_dir + '\\' + filename
    workbook = xlsxwriter.Workbook(directory)
    worksheet = workbook.add_worksheet()

    # Landscape orientation
    worksheet.set_landscape()
    # A4
    worksheet.set_paper(9)
    # Margins
    worksheet.set_margins(left=0.7, right=0.7, top=0.7, bottom=0.7)

    # background colors
    # (0, 0)
    first_elem = '#BF8F00'
    # times and dates
    time_date = '#FFE699'
    # class
    has_text = '#FDE9D8'
    # no class
    no_text = '#F4B084'

    # Initial writings
    for row in cldr.calendar:

        for i in range(len(row)):
            # Cell_format without background color
            cell_format = create_format(workbook)
            # Base color
            cell_format.set_bg_color(time_date)
            # (0, 0)
            if i == 0 and row == cldr.calendar[0]:
                cell_format.set_bg_color(first_elem)
            # No text
            elif i != 0 and row != cldr.calendar[0]:
                cell_format.set_bg_color(no_text)
            # Time
            elif i == 0 and row != cldr.calendar[0]:
                row[i] = row[i] + ':00'

            # Write on Excel
            worksheet.write(cldr.calendar.index(row), i, row[i], cell_format)

    # Fix cell size
    worksheet.set_column(0, len(cldr.calendar[0]), cldr.column_width)
    worksheet.set_default_row(cldr.row_height)

    # Do merging
    for pointA, pointB, text in cldr.merge_cells:
        # Cell format without background color
        cell_format = create_format(workbook)

        # Add bg color based on location
        # (0, 0)
        if pointA[0] == 0 and pointA[1] == 0:
            cell_format.set_bg_color(first_elem)
        # Time
        elif pointA[0] == 0 and pointA[1] != 0:
            cell_format.set_bg_color(time_date)
        # Date
        elif pointA[0] != 0 and pointA[1] == 0:
            cell_format.set_bg_color(time_date)
        # Text
        elif len(text) > 1:
            cell_format.set_bg_color(has_text)
        # No text
        else:
            cell_format.set_bg_color(no_text)

        # Do merge
        worksheet.merge_range(pointA[0], pointA[1], pointB[0], pointB[1], text, cell_format)

    # Finished Excel file
    workbook.close()


def create_format(workbook):
    """Return base cell_format"""

    cell_format = workbook.add_format({'border': 2})
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    cell_format.set_text_wrap()
    cell_format.set_bold()

    return cell_format
