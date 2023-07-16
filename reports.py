from shutil import copy
from time import sleep
from setting import logger, BASE_PATH, PASSWORD, LOGIN, Pattern
from pathlib import Path
from typing import List
import service
from base import Relevantus
from os import path
from shutil import move
from urllib import parse
import pandas as pd


def create(file: Path):
    dir_name = file.parent
    tasks: List[Pattern] = service.read_pattern(file)
    path_file: Path = Path(path.join(dir_name, "Spisok_vygruzki.xlsx"))

    copy(file, path_file)

    rls = Relevantus(BASE_PATH)
    rls.auth(LOGIN, PASSWORD)

    log_list: List[tuple] = []
    for count, task in enumerate(tasks):
        rls.open_url("https://psc.relevantus.org/services/rha")

        try:
            if task.url:
                logger.info("Заполняю: URL")
                rls.update_window_by_xpath(
                    xpath="//input[@id='user-page-url']",
                    text_input=task.url
                )

            if task.request:
                logger.info("Заполняю: Запрос")
                rls.update_window_by_xpath(
                    xpath="//input[@id='search-query']",
                    text_input=task.request
                )

            if task.region:
                logger.info("Заполняю: Регион")
                rls.click_by_xpath(
                    xpath="//span[@aria-labelledby='select2-yandex-xml-api-search-region-container']"
                )
                rls.update_window_by_xpath(
                    xpath="//span[@class='select2-search select2-search--dropdown']/input[@class='select2-search__field']",
                    text_input=task.region
                )
                rls.click_by_xpath(xpath=f"//li[contains(text(), '{task.region}')]")

            if task.type_url:
                logger.info("Заполняю: Учитывать тип URL")
                rls.click_by_xpath(
                    xpath="//div[@class='input-with-valid']/select[@id='search-pages-location']"
                )
                rls.click_by_xpath(
                    xpath=f"//option[contains(text(), '{task.type_url}')]"
                )

            if task.commercialization:
                logger.info("Заполняю: Учитывать коммерческость")
                rls.click_by_xpath(
                    xpath="//div[@class='input-with-valid']/select[@id='search-pages-commerciality']"
                )
                rls.click_by_xpath(
                    xpath=f"//option[contains(text(), '{task.commercialization}')]"
                )

            if task.size_top:
                logger.info("Заполняю: ТОПа для анализа")
                rls.click_by_xpath(
                    xpath="//div[@class='input-with-valid']/select[@id='search-top-size']"
                )
                rls.click_by_xpath(
                    xpath=f"//option[contains(text(), '{task.size_top}')]"
                )

            if task.name_task:
                logger.info("Заполняю: Имя задачи")
                rls.update_window_by_xpath(
                    xpath="//input[@id='service-task-name']",
                    text_input=task.name_task
                )

            sleep(2)

            # Отправляем задачу в работу
            rls.click_by_xpath("//button[@id='btn-submit-rha']")
            logger.info('Задание отправлено в работу')
            sleep(5)

            if rls.check_text_by_xpath("//strong[text()='Превышена суточная квота.']"):
                log_list.append(("Превышена суточная квота.", None, None, None, None, None, None, None))
                break

            rls.wait_element_by_xpath("//a[@id='rha-go-to-result-link']")
            url_task_result = rls.get_attribute_by_xpath("//a[@id='rha-go-to-result-link']", "href")

            log_list.append((url_task_result, None, None, None, None, None, task.url, task.request))

            if (count % 10 == 0) and (count != 0):
                sleep(30)

        except Exception as ex:
            logger.exception(f"Ошибка при формирование отчетов: {ex}")
            log_list.append((task.url, "Отчет не загружен!", None, None, None, None, task.url, task.request))

    sleep(5)
    rls.close_driver()
    service.create_file(log_list, path_file)


def load(file: Path):
    tasks = service.read_load(file)
    log_list: List[tuple] = []

    rls = Relevantus(BASE_PATH)
    rls.auth(LOGIN, PASSWORD)

    for count, task in enumerate(tasks):
        try:
            load_tz_result = recom_relevant_result = analiz_spam_result = top_unigram_result = top_n_gramm_result = None
            rls.open_url(task.link)
            sleep(1)

            check_status = rls.get_text_by_xpath("//td[@id='service-task-result-header-status']").strip()
            if check_status == 'Провалилась':
                log_list.append((
                    task.link,
                    "Задача Провалилась, требуется ручная проверка",
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                ))
                logger.info("Задача Провалилась, требуется ручная проверка")
                continue
            if check_status != 'Выполнена':
                log_list.append((
                    task.link,
                    "Отчет не выгружен, требуется ручная проверка",
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                ))
                logger.info("Отчет не выгружен, требуется ручная проверка")
                continue

            if task.load_tz:
                logger.info("Скачиваем отчет: Выгрузка_ТЗ")
                rls.click_by_xpath(xpath="//button[@id='rha-overall-report-docx-button']")
                name_file = service.wait_file_and_return_path(BASE_PATH, "report_normalized.docx")
                load_tz_result = 'Файл не найден'
                if name_file:
                    DIR_LOAD_TZ = service.make_dir_project(path.join(file.parent, 'Выгрузка_ТЗ'))
                    load_tz_result = path.join(DIR_LOAD_TZ, f"{parse.urlsplit(task.url).path.replace('/', '')}.docx")
                    service.check_busy_file(name_file)
                    service.check_busy_file(load_tz_result)
                    move(name_file, load_tz_result)
                    logger.info(f"Файл {name_file} найден и перенесен")

            if task.recommendations:
                logger.info("Скачиваем отчет: Рекомендации_по_улучшению_релевантности")
                rls.click_by_xpath(xpath="//button[@id='rha-recommendations-xlsx-button']")
                name_file = service.wait_file_and_return_path(BASE_PATH, "recommendations_normalized.xlsx")
                recom_relevant_result = 'Файл не найден'
                if name_file:
                    DIR_RECOM_RELEVANT = service.make_dir_project(
                        path.join(file.parent, 'Рекомендации_по_улучшению_релевантности'))
                    recom_relevant_result = path.join(DIR_RECOM_RELEVANT,
                                                      f"{parse.urlsplit(task.url).path.replace('/', '')}.xlsx")
                    service.check_busy_file(name_file)
                    service.check_busy_file(recom_relevant_result)
                    move(name_file, recom_relevant_result)
                    logger.info(f"Файл {name_file.name} найден и перенесен")

            if task.spam:
                logger.info("Скачиваем отчет: Анализ_переспама")
                rls.click_by_xpath(xpath="//button[@id='rha-max-analysis-xlsx-button']")
                name_file = service.wait_file_and_return_path(BASE_PATH, "max_analysis_normalized.xlsx")
                analiz_spam_result = 'Файл не найден'
                if name_file:
                    DIR_ANALIZ_SPAM = service.make_dir_project(path.join(file.parent, 'Анализ_переспама'))
                    analiz_spam_result = path.join(DIR_ANALIZ_SPAM,
                                                   f"{parse.urlsplit(task.url).path.replace('/', '')}.xlsx")
                    service.check_busy_file(name_file)
                    service.check_busy_file(analiz_spam_result)
                    move(name_file, analiz_spam_result)
                    logger.info(f"Файл {name_file.name} найден и перенесен")

            if task.top_unigrams:
                logger.info("Скачиваем отчет: ТОП_униграмм")
                rls.click_by_xpath(xpath="//button[@id='rha-unigrams-top-xlsx-button']")
                name_file = service.wait_file_and_return_path(BASE_PATH, "unigrams_top_normalized.xlsx")
                top_unigram_result = 'Файл не найден'
                if name_file is not None:
                    DIR_TOP_UNIGRAM = service.make_dir_project(path.join(file.parent, 'ТОП_униграмм'))
                    top_unigram_result = path.join(DIR_TOP_UNIGRAM,
                                                   f"{parse.urlsplit(task.url).path.replace('/', '')}.xlsx")
                    service.check_busy_file(name_file)
                    service.check_busy_file(top_unigram_result)
                    move(name_file, top_unigram_result)
                    logger.info(f"Файл {name_file.name} найден и перенесен")

            if task.top_n_gramm:
                logger.info("Скачиваем отчет: ТОП_N_грамм")
                rls.click_by_xpath(xpath="//button[@id='rha-ngrams-top-xlsx-button']")
                name_file = service.wait_file_and_return_path(BASE_PATH, "ngrams_top_normalized.xlsx")
                top_n_gramm_result = 'Файл не найден'
                if name_file:
                    DIR_TOP_N_GRAMM = service.make_dir_project(path.join(file.parent, 'ТОП_N_грамм'))
                    top_n_gramm_result = path.join(DIR_TOP_N_GRAMM,
                                                   f"{parse.urlsplit(task.url).path.replace('/', '')}.xlsx")
                    service.check_busy_file(name_file)
                    service.check_busy_file(top_n_gramm_result)
                    move(name_file, top_n_gramm_result)
                    logger.info(f"Файл {name_file.name} найден и перенесен")

            text = rls.get_text_by_xpath(
                "//div[contains(. , 'Охват (ширина облака)')]/div[@class='block-below-table']")
            width, depth, relevance = service.get_width_depth_relevance(text)

            if (count % 10 == 0) and (count != 0):
                sleep(30)

            sleep(2)
            logger.info(f"Загрузка {task.url} завершена")
            log_list.append((
                task.link,
                task.url,
                task.request,
                width,
                depth,
                relevance,
                load_tz_result,
                recom_relevant_result,
                analiz_spam_result,
                top_unigram_result,
                top_n_gramm_result,
            ))
        except Exception as ex:
            logger.exception(f"Ошибка при загрузке отчетов: {ex}")

    sleep(5)
    rls.close_driver()
    service.create_result_file(log_list, file)


def analysis(file: Path):
    logger.info("Формирую отчет...")

    RELEVANTUS_DATA_ANALYSIS = path.join(file.parent, 'Relevantus_data_analysis.xlsx')
    result_df = pd.read_excel(file, sheet_name='Результаты')
    analiz_spam = result_df[['url', 'Анализ переспама']].values.tolist()
    recommendations = result_df[['url', 'Рекомендации по улучшению релевантности']].values.tolist()

    with pd.ExcelWriter(RELEVANTUS_DATA_ANALYSIS, engine='xlsxwriter') as writer:
        service.re_spam(analiz_spam).to_excel(writer, sheet_name='Переспам', index=False)
        service.replay_word(recommendations).to_excel(writer, sheet_name='Повторы слов', index=False)
        service.add_common_word(recommendations).to_excel(writer, sheet_name='Добавить важные слова', index=False)
        service.dop_word(recommendations).to_excel(writer, sheet_name='Доп. слова', index=False)
        service.title(recommendations).to_excel(writer, sheet_name='Title', index=False)

    service.wight_row(RELEVANTUS_DATA_ANALYSIS)

    logger.info(f'Отчет сформирован, доступен по следующему пути: \n{RELEVANTUS_DATA_ANALYSIS}')
