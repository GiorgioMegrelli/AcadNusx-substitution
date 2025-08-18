from abc import ABC
from typing import List, Optional, Tuple


class IFilter(ABC):
    def filter(self, val: Optional[str]) -> bool:
        raise NotImplementedError()


class AllTrueFilter(IFilter):
    def filter(self, _: Optional[str]) -> bool:
        return True


class FontNameFilter(IFilter):
    def __init__(self, font: str | Tuple[str] | List[str]):
        if isinstance(font, str):
            self._font = {font.lower()}
        elif isinstance(font, (list, tuple, set)):
            self._font = {f.lower() for f in font}
        else:
            raise ValueError(f"Invalid format: {font}")

    def filter(self, val: Optional[str]) -> bool:
        if val is None:
            return False
        return val.lower() in self._font
