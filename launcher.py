import time
import schedule

from page_parser import parse_page

def launch():

    l_time = time.localtime()
    human_readable_date = time.strftime("%d.%m.%y", l_time)

    parse_page(f'scrape-{human_readable_date}')

if __name__ == '__main__':

    schedule.every().week.do(launch)

    while True:

        schedule.run_pending()
        time.sleep(1)