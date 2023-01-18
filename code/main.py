from services import scraper_service, data_processing_service, excel_service
from constants import TIME_PERIOD

def main():
    period = int(input('Введите количество дней (default = 30): ') or TIME_PERIOD)
    head, data = scraper_service.scrap(period)
    filename = data_processing_service.data_processing(head, data)
    excel_service.create_pivot_table(filename)

if __name__ == '__main__':
    main()
