from   dotenv import load_dotenv
from   typing   import Literal
from   datetime import datetime
import pandas   as pd
import logging
import colorlog
import copy
import time
import sys
import os


load_dotenv()
os.chdir(os.getenv('ROOT_DIR'))

handler = colorlog.StreamHandler(sys.stdout)
handler.setFormatter(colorlog.ColoredFormatter(
    '%(log_color)s%(message)s',
    log_colors={
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'bold_red',
    }
))

logger = logging.getLogger(__name__)
logger.addHandler(handler)
logger.setLevel(logging.INFO)
logger.propagate = False



class Object:

    _dpath = os.getenv('INPUT_')
    _epath = os.getenv('OUTPUT_')

    def __init__(self, pth, sep = ';', s_name = 0, mpath: str = None, epath: str = None):
        self.path   =  pth
        self.s_name =  s_name
        self.epath  =  epath or self._epath
        self.mpath  =  mpath or self._dpath
        self.data   =  self.to_df(pth, sep)
        self.name   =  f"{os.path.splitext(pth)[0]}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"


    def to_df(self, filepath: str, sep: str = ';') -> pd.DataFrame | None:
        ext = os.path.splitext(filepath)[1].lower()
        full_path = os.path.join(self.mpath, filepath)
        try:
            if ext == '.xlsx':
                with pd.ExcelFile(full_path, engine='openpyxl') as xls:
                    sheets = xls.sheet_names
                    selected = (
                        sheets[self.s_name] if isinstance(self.s_name, int) else self.s_name
                    )
                    df = pd.read_excel(xls, sheet_name=selected)

                logger.info(
                    "Файл '%s' \nТаблицы: %s, \nВыбрана '%s' (s_name=%s),\n Shape=%s\n",
                    filepath, sheets, selected, self.s_name, df.shape
                )
                return df

            elif ext in ('.csv', '.txt'):
                df = pd.read_csv(full_path, sep=sep)
                logger.info(
                    "Файл '%s' открылся с разделителем '%s' shape = %s\n",
                    filepath, sep, df.shape
                )
                return df
            else:
                logger.warning('Неподдерживаемое расширение файла: %s (%s)', ext, filepath)
                return None
        except Exception as ex:
            logger.error('Ошибка при чтении %s: %s', filepath, ex)
            return None


    def _ensure_data(self, action: str) -> bool:
        if self.data is None:
            logger.error('%s невозможно: данные не загружены', action)
            return False
        return True

    def cast(self, column: str | list[str], dtype: str | type | list) -> 'Object':
        if not self._ensure_data('Приведение типа'):
            return self

        columns = column if isinstance(column, list) else [column]
        dtypes = dtype if isinstance(dtype, list) else [dtype] * len(columns)
        if len(columns) != len(dtypes):
            logger.warning(
                "Количество колонок и типов не совпадает: %s колонок, %s типов — приведение отменено",
                len(columns), len(dtypes)
            )
            return self

        for col, dt in zip(columns, dtypes):
            if col not in self.data.columns:
                logger.warning("Колонка '%s' не найдена — приведение пропущено", col)
                continue
            try:
                pd.api.types.pandas_dtype(dt)
            except TypeError:
                logger.warning("Неподдерживаемый тип '%s' для колонки '%s' — пропущено", dt, col)
                continue
            self.data[col] = self.data[col].astype(dt)
        return self

    def rename(self, col: str | list[str], new_name: str | list[str]) -> 'Object':
        if not self._ensure_data('Переименование'):
            return self

        columns = col if isinstance(col, list) else [col]
        new_names = new_name if isinstance(new_name, list) else [new_name]

        if len(columns) != len(new_names):
            logger.warning(
                "Разное количество передаваемых столбцов и имён, возможен пропуск:\n%s -- %s ; %s -- %s\n"
                , columns, len(columns), new_names, len(new_names)
            )

        for col, name in zip(columns, new_names):
            if col not in self.data.columns:
                logger.warning("Колонка '%s' не найдена — переименование не выполнено", col)
                continue
            self.data = self.data.rename(columns={col: name})
        return self

    def drop(self, column: list[str] | str) -> 'Object':
        if not self._ensure_data('Удаление колонок'):
            return self
        cols = [column] if isinstance(column, str) else column
        missing = [c for c in cols if c not in self.data.columns]
        if missing:
            logger.warning("Колонки не найдены и будут проигнорированы: %s", missing)
        self.data = self.data.drop(columns=cols, errors='ignore')
        return self


    def fillna(self,
               column: str | list[str],
               symbol: str | list[str]) -> 'Object':

        if not self._ensure_data('Заполнение пропусков'):
            return self

        columns = column if isinstance(column, list) else [column]
        symbols = symbol if isinstance(symbol, list) else [symbol]

        if len(symbols) == 1:
            symbols = symbols * len(columns)

        elif len(symbols) != len(columns):
            logger.warning(
                "Количество колонок и значений замены не совпадает: %s колонок, %s значений — заполнение отменено\n",
                len(columns), len(symbols)
            )
            return self

        missing = [c for c in columns if c not in self.data.columns]
        if missing:
            logger.warning("Колонки не найдены и будут пропущены: %s\n", missing)

        fill_map = {c: s for c, s in zip(columns, symbols) if c not in missing}
        self.data = self.data.fillna(fill_map)
        return self

    def replace(self,
                column:  str | list[str],
                symbol:  str | list[str|list[str]],
                replace: str | list[str|list[str]] ) -> 'Object':
        if not self._ensure_data('Замена символов\n'):
            return self

        columns  = column  if isinstance(column, list)  else [column]
        symbols  = symbol  if isinstance(symbol, list)  else [symbol]
        replaces = replace if isinstance(replace,list)  else [replace]

        if not(len(columns) == len(symbols) == len(replaces)):
            logger.error(
                "Количество колонок и значений замены не совпадает:\n"
                     "Колонки: \n%s,\n Заменяемых символов : \n%s,\n Замен %s\n"
                     "Отмена операции\n",
                      columns, symbols, replaces
            )
            return self

        replace_dict = {}

        for col, sym, rep in zip(columns, symbols, replaces):

            if col not in self.data.columns:
                logger.warning(
                    "Колонка %s не присутствует в таблице и будет пропущена",
                    col
                )
                continue

            if  isinstance(sym, str)  and  isinstance(rep, str):
                replace_dict[col] = {sym : rep}

            elif isinstance(sym, list) and  isinstance(rep, list):
                if len(sym) != len(rep):
                    logger.warning(
                        "Кол-во символов и их замен не совпадает - возможны пропуск\n"
                        "Заменяемые символы - %s\nЗамены - %s\n",
                        sym,rep)
                replace_dict[col] = {s: r for s, r in zip(sym,rep)}
            else:
                logger.warning(
                    "Неожиданное сочетание типов для колонки '%s': symbol=%s, replace=%s — пропущено",
                    col, type(sym), type(rep)
                )

        self.data = self.data.replace(replace_dict)
        logger.info('Проведённые замены:\n%s\n', replace_dict)
        return self

    def drop_duplicates(self,
                        subset: list[str] | str = None,
                        keep: Literal['first', 'last', False] = 'first') -> 'Object':

        if not self._ensure_data('Удаление дубликатов'):
            return self

        before = self.data.shape[0]
        self.data = self.data.drop_duplicates(subset=subset, keep=keep)
        after = self.data.shape[0]

        if before != after:
            logger.info("Удалено дубликатов: %s (было %s строк, стало %s)", before - after, before, after)
        else:
            logger.info("Дубликатов не найдено")
        return self


    def export(self, name: str = None, index: bool = False):
        if not self._ensure_data('Экспорт'):
            return self
        name = name or self.name
        self.data.to_excel(os.path.join(self.epath, name) + '.xlsx', index=index)
        os.startfile(self.epath)
        return self

    def join(self,
             joined: 'Object',
             on: str | list = None,
             how: Literal['left', 'right', 'inner', 'outer', 'cross'] = 'left',
             left_on: str | list = None,
             right_on: str | list = None,
             inplace: bool = False) -> 'Object | None':

        if not self._ensure_data(f'Join') or not joined._ensure_data(f'Join с {joined.name}'):
            return self if not inplace else None

        def _as_list(x):
            return x if isinstance(x, list) else [x]

        if on:
            missing = [c for c in _as_list(on) if c not in self.data.columns or c not in joined.data.columns]
            if missing:
                logger.error(
                    "Колонки %s отсутствуют в одной из таблиц\nОтмена операции",
                    missing
                )
                return self

            for col in _as_list(on):
                if self.data[col].dtype != joined.data[col].dtype:
                    logger.warning(
                        "Разный тип данных столбца '%s': initial - %s   joined - %s",
                        col, self.data[col].dtype, joined.data[col].dtype
                    )

        elif left_on and right_on:
            missing_left = [c for c in _as_list(left_on) if c not in self.data.columns]
            missing_right = [c for c in _as_list(right_on) if c not in joined.data.columns]
            if missing_left or missing_right:
                logger.error(
                    "Один из аргументов left_on = %s или right_on = %s не присутствует в одной из таблиц\nОтмена операции",
                    left_on, right_on
                )
                return self

            for l_col, r_col in zip(_as_list(left_on), _as_list(right_on)):
                if self.data[l_col].dtype != joined.data[r_col].dtype:
                    logger.warning(
                        "Разный тип данных столбца: initial['%s'] - %s   joined['%s'] - %s",
                        l_col, self.data[l_col].dtype, r_col, joined.data[r_col].dtype
                    )

        merged = self.data.merge(joined.data, how=how, on=on, left_on=left_on, right_on=right_on)

        if inplace:
            self.data = merged
            return self
        else:
            result = copy.copy(self)  # или через _from_data, как обсуждали ранее
            result.data = merged
            result.name = f"{self.name}_join_{joined.name}"
            result.path = ''
            return result

    @property
    def shape(self):
        return self.data.shape if self.data is not None else None

    @property
    def col(self):
        return self.data.columns if self.data is not None else None

    @property
    def dtypes(self):
        return self.data.dtypes if self.data is not None else None

    @classmethod
    def concat(
            cls,
            *objects: 'Object',
            axis: Literal[0, 1] = 0) -> 'Object | None':

        valid = [obj for obj in objects if obj._ensure_data(f'Concat: {obj.name}')]

        if len(valid) < 2:
            logger.warning(
                "Недостаточно валидных объектов для concat: %s из %s — операция отменена",
                len(valid), len(objects)
            )
            return None

        if len(valid) < len(objects):
            logger.warning(
                "Пропущено %s объект(ов) с пустыми данными при concat",
                len(objects) - len(valid)
            )

        merged = pd.concat([obj.data for obj in valid], axis=axis, ignore_index=True)

        result = copy.copy(valid[0])
        result.data = merged
        result.name = f"concat_{len(valid)}_frames_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
        result.path = ''

        logger.info(
            "Concat выполнен: %s объектов, итоговый shape=%s",
            len(valid), result.shape
        )
        return result








