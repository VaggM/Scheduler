from html_info.html_explore import get_html_lessons
from gui.my_window import MyWindow


def get_name(lesson_inst):
    """Gets the name of a lesson"""
    return lesson_inst.info['course_name']


# Get all lessons from html
filename = 'url.htm'
url_lessons = get_html_lessons(filename)
url_lessons.sort(key=get_name)

# Find all unique names
uniques = []
for lesson in url_lessons:
    if get_name(lesson) not in uniques:
        uniques.append(get_name(lesson))

# Make a window with a list of unique names
#   and give it all the lessons you found within the html
#   for Excel usage
window = MyWindow(uniques, url_lessons)
window.main()
