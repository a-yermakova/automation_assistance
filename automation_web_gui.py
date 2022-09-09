""" Module defining web graphic user interface for Automation assistance. """

import streamlit as st
from io import BytesIO

from automation_assistance import (get_sheetnames_with_binary_stream,
                                   tags_equations_creator,
                                   check_cell_address_input,
                                   add_similar_statement_tags_from_config,
                                   specify_function_for_cell_address_searching)
from automation_assistance_exceptions import EmptyTagCellInModel


def change_app_status(new_status: str = None):
    """Функция для изменения статуса сессии приложения.
    Args:
        :param new_status: значение статуса приложения, может быть "before", "on", "after"
        :type new_status: str
    """
    if new_status:
        st.session_state.status = new_status
    match st.session_state.status:
        case "before":
            page_before_updating()
        case "on":
            page_on_updating()
        case "after":
            page_after_updating()

def page_before_updating():
    """Функция GUI для отображения начального окна приложения, в котором пользователь загружает
    файлы формата XLSX для дальнейшей обработки"""
    with column_one:
        st.session_state.model_file = st.file_uploader('Загрузите файл аналитической модели',
                                                       type=source, accept_multiple_files=False)
        if st.session_state.model_file:
            st.session_state.model_binary_stream = BytesIO(st.session_state.model_file.getvalue())
            #получаем названия страниц загруженного XLSX файла
            st.session_state.model_sheet_names = get_sheetnames_with_binary_stream(st.session_state.model_binary_stream)
            st.session_state.selected_model_sheet = st.selectbox('Выберите название листа для сравнения в файле аналитической модели',
                                                                 st.session_state.model_sheet_names)
    with column_two:
        st.session_state.issuer_file = st.file_uploader('Загрузите файл эмитента',
                                                        type=source, accept_multiple_files=False)
        if st.session_state.issuer_file:
            st.session_state.issuer_binary_stream = BytesIO(st.session_state.issuer_file.getvalue())
            # получаем названия страниц загруженного XLSX файла
            st.session_state.issuer_sheet_names = get_sheetnames_with_binary_stream(
                                                            st.session_state.issuer_binary_stream)
            st.session_state.selected_issuer_sheet = st.selectbox(
                                                        'Выберите название листа для сравнения в файле эмитента',
                                                        st.session_state.issuer_sheet_names)
    if st.button('Отправить на обработку', key='fileprocessing'):
        if st.session_state.model_file and st.session_state.issuer_file:
            with st.spinner('Обработка...'):
                st.session_state.status = 'on'
                st.experimental_rerun()
        else:
            st.error('Загрузите файлы', icon="🚨")

def page_on_updating():
    """Функция GUI для отображения окна после загрузки файлов:
    пользователь выбирает страницы из двух файлов, которые необходимо обработать,
    и вводит адреса верхних ячеек столбцов с числовыми значениями.
    Происходит обработка с помощью сторонних функций."""
    model_binary_stream = st.session_state.model_binary_stream
    issuer_binary_stream = st.session_state.issuer_binary_stream
    selected_model_sheet = st.session_state.selected_model_sheet
    selected_issuer_sheet = st.session_state.selected_issuer_sheet
    default_model_address_of_start, default_issuer_address_of_start = specify_function_for_cell_address_searching(
                                                                        model_binary_stream,
                                                                        issuer_binary_stream,
                                                                        selected_model_sheet,
                                                                        selected_issuer_sheet)
    with column_one:
        model_address_of_start = st.text_input(label='Введите адрес верхней ячейки столбца с числовыми значениями\n'
                                               + 'в файле аналитической модели (в формате А1):',
                                               value=default_model_address_of_start)
        statement_block = st.selectbox('Укажите блок статей', statement_block_option)
    with column_two:
        issuer_address_of_start = st.text_input(label='Введите адрес верхней ячейки столбца с числовыми значениями\n'
                                                + 'в файле эмитента (в формате А1):',
                                                value=default_issuer_address_of_start)
        data_source = st.selectbox('Укажите источник данных', data_source_option)

    if st.button('Отправить на обработку', key='sendtoprocessing'):
        # Если ввели адреса ячеек
        if model_address_of_start and issuer_address_of_start:
            # Если формат адреса ячеек верный
            model_address_of_start = check_cell_address_input(model_address_of_start)
            issuer_address_of_start = check_cell_address_input(issuer_address_of_start)
            if model_address_of_start and issuer_address_of_start:
                with st.spinner('Обработка...'):
                    # сохраняем список названий статей из модели и соответствующих им названий из файла эмитента
                    try:
                        st.session_state.list_of_equivalents = tags_equations_creator(
                                                                      model_binary_stream=model_binary_stream,
                                                                      issuer_binary_stream=issuer_binary_stream,
                                                                      selected_model_sheet=selected_model_sheet,
                                                                      selected_issuer_sheet=selected_issuer_sheet,
                                                                      model_address_of_start=model_address_of_start,
                                                                      issuer_address_of_start=issuer_address_of_start
                                                                      )
                        # добавляем названия статей из других аналитических моделей (из конфига),
                        # если в соответствующих им статьях (справа в конфиге) есть совпадение
                        # с переданным названием статьи у эмитента
                        add_similar_statement_tags_from_config(statement_block, data_source,
                                                               st.session_state.list_of_equivalents)
                        st.session_state.status = 'after'
                        st.experimental_rerun()
                    except EmptyTagCellInModel as error:
                        st.error(error, icon='🚨')

            else:
                st.error('Неверный формат записи ячейки', icon='🚨')
        else:
            st.error('Введите значения', icon='🚨')

def page_after_updating():
    """Функция GUI для отображения итогового окна приложения, в котором пользователь может
    просмотреть список полученных эквивалентных значений,
    а также скачать его в формате .txt."""

    output_text = ''
    for equivalent in st.session_state.list_of_equivalents:
        if len(equivalent.tags_from_config) == 1 and equivalent.model_tag == equivalent.tags_from_config[0]:
            equal_statements = f'{equivalent.model_tag} = {equivalent.tag_from_filling}\n' \
                               f'☻ Переименование не требуется!\n'
            output_text += equal_statements + '\n'
        else:
            equal_statements = f'{equivalent.model_tag} = {equivalent.tag_from_filling}\n' \
                               f'► Переименовать в: \n'
            statements_to_rename = '\n'.join(equivalent.tags_from_config)
            output_text += equal_statements + statements_to_rename + '\n\n'
    st.text_area('Полученный список:', output_text, height=500)
    st.download_button(
        label='Скачать в виде файла',
        data=output_text,
        file_name='list_of_equivalents.txt',
        mime='text'
        )
    if st.button('Начать заново', key='restart'):
        st.session_state.status = 'before'
        st.experimental_rerun()

# ___________________________________________________________________________________________________________________
source = 'xlsx'
statement_block_option = ['Баланс', 'Финансовые результаты', 'Сегменты', 'Отчет о движении денежных средств']
data_source_option = ['XLSX', 'PDF', 'XBRL']

st.set_page_config(page_title='Automation assistance',
                   page_icon='👩‍🎓',
                   layout="wide",
                   initial_sidebar_state="collapsed"
                   )
st.header('Automation assistance', )
column_one, column_two = st.columns((1, 1))
if 'status' not in st.session_state:
    st.session_state.update({'list_of_equivalents': []})
    change_app_status('before')
else:
    change_app_status()

#python -m streamlit run .\automation_web_gui.py
#python -m streamlit run .\��_fill_in.py