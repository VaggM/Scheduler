from bs4 import BeautifulSoup
from html_info.lesson import Lesson


def get_html_lessons(filename):
    """Returns all lessons found in events area of the html"""
    # Get html text
    with open(filename, 'rb') as fp:
        soup = BeautifulSoup(fp, "html.parser")

    text = soup.get_text()

    # Find events area in html file
    # Starting string
    event_start = text.find('events : [')
    text = text[event_start:]
    # Ending string
    event_end = text.find('});')
    text = text[:event_end]

    # Start saving lessons
    lessons = []

    # Find all text starting with 'id:', they define a lesson event
    while True:
        # Finds first available 'id: '
        first_id = text.find('id:')
        # tmp string to find the second 'id:'
        tmp = text[first_id + 1:]
        second_id = tmp.find('id:')
        # Make a lesson based on text between first and second 'id:'
        lesson_txt = text[first_id:second_id]
        lessons.append(Lesson(lesson_txt))
        # Remove all used text so far
        text = text[second_id:]
        # Final lesson, no more 'id:'
        if second_id == -1:
            break
    # Return all found lessons
    return lessons
