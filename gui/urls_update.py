import requests


def urls_update(current_urls):

    url = "https://github.com/VaggM/Scheduler/tree/main/urls"

    resp = requests.get(url)

    txt = resp.text

    urls = []

    while True:

        next_spring = txt.find('spring')
        next_winter = txt.find('winter')

        winter = next_winter != -1
        spring = next_spring != -1

        if not winter and not spring:
            break
        elif winter and (next_winter < next_spring or not spring):
            trial = txt[next_winter-30:next_winter+30]
            title = trial.find('title="')
            if title != -1:
                title = trial[title+7:]
                end = title.find('"')
                title = title[:end]
                urls.append(title)
            txt = txt[next_winter+1:]
        elif spring and (next_spring < next_winter or not winter):
            trial = txt[next_spring-30:next_spring+30]
            title = trial.find('title="')
            if title != -1:
                title = trial[title + 7:]
                end = title.find('"')
                title = title[:end]
                urls.append(title)
            txt = txt[next_spring + 1:]

    for url in urls:

        text_url = "https://raw.githubusercontent.com/VaggM/Scheduler/main/urls/" + url
        resp = requests.get(text_url)
        text = resp.text
        with open(f"urls\\{url}", "w") as f:
            f.write(text)
