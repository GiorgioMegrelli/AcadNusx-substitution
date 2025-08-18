import click


class UnsupportedWordDocFormat(click.BadParameter):
    def __init__(self, file_ext: str, value: str):
        msg = f"Unsupported Word Document format '.{file_ext}' for '{value}'"
        super().__init__(msg)


class BadFileExtension(click.BadParameter):
    def __init__(self, file_ext: str, value: str):
        msg = f"Bad file format '.{file_ext}' for '{value}'"
        super().__init__(msg)
