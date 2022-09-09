"""Module for keeping custom exceptions for automation assistance."""


class EmptyTagCellInModel(Exception):
    """Ошибка пустой строки в тэгах модели, когда найденное значение за сопоставимые периоды совпадает."""

