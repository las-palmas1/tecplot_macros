import typing
import os
import pandas as pd
import copy
import enum
import numpy as np


class Point:
    def __init__(self, x, y, z):
        self.x = x
        self.y = y
        self.z = z


class PolyLine:
    def __init__(self, nodes: typing.List[Point], numpoints):
        self.nodes = nodes
        self.numpoints = numpoints


def wrap_macro(macro_body: str) -> str:
    result = "#!MC 1410\n" \
             "$!VarSet |MFBD| = 'C:\Program Files\Tecplot\Tecplot 360 EX 2016 R2'\n" \
             "%s"\
             "$!RemoveVar |MFBD|\n" \
             "$!Quit" % macro_body
    return result


def _get_extract_from_polyline_command(polyline: PolyLine,
                                       filename) -> str:
    result = "$!EXTRACTFROMPOLYLINE\n" \
           "EXTRACTLINEPOINTSONLY = NO\n" \
           "EXTRACTTHROUGHVOLUME = YES\n" \
           "EXTRACTTOFILE = YES\n" \
           "FNAME = '%s'\n" \
           "NUMPTS = %s\n" \
           "RAWDATA\n" \
           "%s\n" % (filename, polyline.numpoints, len(polyline.nodes))
    template = "%s %s %s\n"
    for i in polyline.nodes:
        result += template % (i.x, i.y, i.z)
    return result


def get_open_data_file_command(filename: str):
    """
    Фукция возвращает строку, содержащую набор команд на скриптовом языке TecPlot для 
    открытия файла с расширением .plt\n
    :param filename: str \n 
        Путь к файлу с расширенем .plt
    :return: str \n 
    """
    result = "$!READDATASET  '%s'\n" % filename
    return result


def get_open_layout_command(filename: str) -> str:
    """
    Фукция возвращает строку, содержащую набор команд на скриптовом языке TecPlot для 
    открытия файла с расширением .lay\n
    :param filename: str \n 
        Путь к файлу с расширенем .lay
    :return: str \n 
    """
    result = "$!OPENLAYOUT  '%s'\n" % filename
    return result


def get_save_layout_command(filename: str) -> str:
    """
    Фукция возвращает строку, содержащую набор команд на скриптовом языке TecPlot для 
    сохранения файла с расширением .lay или .plt\n
    :param filename: str \n 
        Путь к файлу с расширенем .lay
    :return: str \n 
    """
    result = "$!SAVELAYOUT  '%s'\n" \
             "  INCLUDEDATA = YES\n" \
             "  INCLUDEPREVIEW = NO\n" % filename
    return result


def _get_data_file_extraction_macro_body(filename, polylines: typing.List[PolyLine], output_dir) -> str:
    s1 = get_open_data_file_command(filename)
    s2 = ''
    output_filename_template = os.path.splitext(os.path.split(filename)[1])[0] + '_line_%s.dat'
    for n, polyline in enumerate(polylines):
        s2 += '%s' % (_get_extract_from_polyline_command(polyline,
                                                         os.path.join(output_dir, output_filename_template) % n))
    return s1 + s2


def create_macro_file(macros: str, filename: str):
    """
    :param macros: str \n
        Строка, содержащая набор команд на скриптовом языке TecPlot
    :param filename: str \n
        Имя файла, в который будет сохранен макрос
    :return: 
    """
    file = open(filename, 'w')
    file.write(macros)
    file.close()


def execute_macro(filename):
    os.system(filename)


class LineDataExtractor:
    """
    Обеспечивает возможность создания макроса, извелающего из .plt файла данные по набору полилиний и
    сохраняющего эти данные в текстовые файлы с расширением .dat
    """
    def __init__(self, datafiles_dir, output_dir, polylines_list: typing.List[typing.List[PolyLine]], macro_name):
        """
        :param datafiles_dir: str \n
            Имя директории, в которой располагаются .plt или файлы
        :param output_dir: str \n
            Имя директории для извлеченных данных
        :param polylines_list: List[List[PolyLine]] \n
            список полилиний, по которым будут извелкаться данные
        :param macro_name: str \n
            имя макроса, под которым он будет сохранен
        """
        self.datafiles_dir = datafiles_dir
        self.output_dir = output_dir
        self.polylines_list = polylines_list
        self.macro_name = macro_name

    def _get_macro(self) -> str:
        data_filenames = os.listdir(self.datafiles_dir)
        assert len(self.polylines_list) == len(data_filenames), 'Number of data files and number of sets of polylines '\
                                                                'must be same'
        macros_body = ''
        for n, filename in enumerate(data_filenames):
            if os.path.splitext(filename)[1] == '.plt':
                macros_body += '%s' % _get_data_file_extraction_macro_body(os.path.join(self.datafiles_dir, filename),
                                                                           self.polylines_list[n], self.output_dir)
        return wrap_macro(macros_body)

    def run_extraction(self):
        macro = self._get_macro()
        create_macro_file(macro, self.macro_name)
        execute_macro(self.macro_name)


class LineDataLoader:
    """
    Позволяет считать извлеченные по полилиниям данные  из всех файлов, содержащихся в папке, путь к которой 
    хранится в поле data_dirname, и записать данные из каждого файла в экземпляр класса Pandas.DataFrame, доступ к 
    списку которых осуществляется через поле frame
    """
    def __init__(self, data_dirname: str):
        """
        :param data_dirname: str \n
            имя папки, содержащей файлы с извлеченными данными
        """
        self.data_dirname = data_dirname
        self.frames: typing.List[pd.DataFrame] = []

    @classmethod
    def _split_str(cls, string: str) -> typing.List[str]:
        arr = string.split(sep=' ')
        res = []
        for s in arr:
            if s != '' and s != '\n':
                res.append(s)
        return res

    @classmethod
    def _format_strings(cls, str_list: typing.List[str]) -> typing.List[str]:
        result = copy.deepcopy(str_list)
        for n, i in enumerate(str_list):
            result[n] = i.replace('\n', '').replace('"', '')
        return result

    @classmethod
    def _get_str_list_from_file(cls, filename) -> typing.List[typing.List[str]]:
        result = []
        file = open(filename, 'r')
        s = ' '
        while s != '':
            s = file.readline()
            result.append(cls._split_str(s))
        file.close()
        for n in range(len(result)):
            result[n] = cls._format_strings(result[n])
        return result

    @classmethod
    def _get_sum_str_list(cls, str_list: typing.List[str], start: int, with_spaces=False) -> str:
        result = ''
        for n, string in enumerate(str_list):
            if n >= start:
                if with_spaces:
                    result += string + ' '
                else:
                    result += string
        if with_spaces:
            result = result[0:(len(result) - 1)]
        return result

    @classmethod
    def _get_variable_names(cls, str_list: typing.List[typing.List[str]]) -> typing.List[str]:
        result = []
        for n, i in enumerate(str_list):
            if n == 0:
                result.append(cls._get_sum_str_list(i, 2, with_spaces=True))
            elif i[0] != 'ZONE':
                result.append(cls._get_sum_str_list(i, 0, with_spaces=True))
            elif i[0] == 'ZONE':
                break
        return result

    @classmethod
    def _get_variable_arrays(cls, str_list: typing.List[typing.List[str]]) -> typing.List[typing.List[float]]:
        result = []
        for n, i in enumerate(str_list):
            try:
                if len(i) > 0:
                    float(i[0])
                    result.append([])
                    for j in i:
                        result[len(result) - 1].append(float(j))
            except ValueError:
                pass
        return result

    def load(self):
        files_list = os.listdir(self.data_dirname)
        for filename in files_list:
            str_list = self._get_str_list_from_file(os.path.join(self.data_dirname, filename))
            var_names = self._get_variable_names(str_list)
            var_arrays = self._get_variable_arrays(str_list)
            self.frames.append(pd.DataFrame.from_records(var_arrays, columns=var_names))


class SliceType(enum.Enum):
    XPLANES = 0
    YPLANES = 1
    ZPLANES = 2
    IPLANES = 3
    JPLANES = 4
    KPLANES = 5
    ARBITRARY = 6


def _get_slice_setting_macro(slice_type: SliceType, position: tuple, **kwargs) -> str:
    position_template = '%s = %s %s = %s %s = %s'
    primary_position = ''
    if slice_type == SliceType.XPLANES or slice_type == SliceType.YPLANES or slice_type == SliceType.ZPLANES or \
       slice_type == SliceType.ARBITRARY:
        primary_position = position_template % ('X', position[0], 'Y', position[1], 'Z', position[2])
    elif slice_type == SliceType.IPLANES or slice_type == SliceType.JPLANES or slice_type == SliceType.KPLANES:
        primary_position = position_template % ('I', position[0], 'J', position[1], 'K', position[2])

    if slice_type == SliceType.ARBITRARY and 'normal' not in kwargs:
        assert 'normal' in kwargs, 'Normal vector must be specified'
    normal_string = ''
    if slice_type == SliceType.ARBITRARY and 'normal' in kwargs:
        normal_string = '$!SLICEATTRIBUTES 1 NORMAl {X = %s Y = %s Z = %s}\n' % (kwargs['normal'][0],
                                                                                 kwargs['normal'][1],
                                                                                 kwargs['normal'][2])
    result = "$!SLICELAYERS SHOW = YES\n" \
             "$!SLICEATTRIBUTES 1  SLICESURFACE = %s\n" \
             "$!SLICEATTRIBUTES 1  PRIMARYPOSITION{%s}\n" % (slice_type.name, primary_position) + normal_string
    return result


class DataType(enum.Enum):
    SINGLE = 0
    SHORTINT = 1
    DOUBLE = 2
    BYTE = 3
    LONGINT = 4
    BIT = 5


def get_alterdata_command(equation: str, ignored_divided_by_zero=False, data_type: DataType = DataType.SINGLE) -> str:
    """
    Возвращает строку, содержащую набор команд на скриптовом языке TecPlot
    :param equation: str \n
        Строка, содержащее уравнение для TecPlot
    :param ignored_divided_by_zero: bool \n
    :param data_type: DataType \n
    :return: str
    """
    template = "$!ALTERDATA\n" \
               "  EQUATION = '%s'\n" \
               "  IGNOREDIVIDEBYZERO = %s\n" \
               "  DATATYPE = %s\n"
    if ignored_divided_by_zero:
        result = template % (equation, 'YES', data_type.name)
    else:
        result = template % (equation, 'NO', data_type.name)
    return result


def _get_go_to_2d_macro(x_axis_var: int, y_axis_var: int, x_line_pos: float=0., y_line_pos: float=0.,
                        rect: tuple = (10, 10, 90, 90), preserve_axis_length: bool = False, **kwargs) -> str:
    """
    :param x_axis_var: int \n
        Номер переменной, откладываемой по горизонтальной оси
    :param x_line_pos: float \n
        Позиция горизонтальной оси по вертикали
    :param y_line_pos: float \n
        Позиция вертикальной оси по горизонтали \n
    :param y_axis_var: int \n
        Номер переменной, откладываемой по вертикальной оси
    :param rect: tuple, optional \n
        Определяет положение прямоугольника сетки на frame, rect=(x1, y1, x2, y2),
        по умолчанию rect=(10, 10, 90, 90)
    :param preserve_axis_length: bool, optional \n
        Сохраняемость масштаба осей при изменении их диапазона
    :param kwargs: xlim, ylim (интервалы по осям x и y соотвественно), тип tuple; пример: xlim=(0,1), ylim=(1,2)
    :return:
    """
    string1 = "$!PLOTTYPE = CARTESIAN2D\n" \
              "$!TWODAXIS XDETAIL{VARNUM = %s}\n" \
              "$!TWODAXIS YDETAIL{VARNUM = %s}\n" % (x_axis_var, y_axis_var)
    string3 = "$!TWODAXIS\n" \
              "  GRIDAREA\n" \
              "  {\n" \
              "    EXTENTS\n" \
              "    {\n" \
              "      X1 = %s\n" \
              "      Y1 = %s\n" \
              "      X2 = %s\n" \
              "      Y2 = %s\n" \
              "    }\n" \
              "  }\n" \
              "$!TWODAXIS XDETAIL{AXISLINE{POSITION = %s}}\n" \
              "$!TWODAXIS YDETAIL{AXISLINE{POSITION = %s}}\n" % (rect[0], rect[1], rect[2], rect[3],
                                                                 x_line_pos, y_line_pos)
    if 'xlim' in kwargs and 'ylim' in kwargs:
        string2 = "$!TWODAXIS\n" \
                  "  PRESERVEAXISSCALE = %s\n" \
                  "  XDETAIL\n" \
                  "    {\n"\
                  "    RANGEMIN = %s\n" \
                  "    RANGEMAX = %s\n" \
                  "    }\n" \
                  "  YDETAIL\n" \
                  "    {\n" \
                  "    RANGEMIN = %s\n" \
                  "    RANGEMAX = %s\n" \
                  "    }\n" % ((not preserve_axis_length).__str__().upper(), kwargs['xlim'][0], kwargs['xlim'][1],
                               kwargs['ylim'][0], kwargs['ylim'][1])
    else:
        string2 = ''
    result = string1 + string2 + string3
    return result


class Font:
    def __init__(self, font_family='Helvetica', is_bold=False, is_italic=False, height=3.):
        self.font_family = font_family
        if is_bold:
            self.is_bold = 'YES'
        else:
            self.is_bold = 'NO'
        if is_italic:
            self.is_italic = 'YES'
        else:
            self.is_italic = 'NO'
        self.height = height


def _get_legend_font_settings(header_font: Font = Font(), number_font: Font = Font()) -> str:
    header_settings = "$!GLOBALCONTOUR 1  LEGEND{HEADERTEXTSHAPE{FONTFAMILY = '%s'}}\n" \
                      "$!GLOBALCONTOUR 1  LEGEND{HEADERTEXTSHAPE{HEIGHT = %s}}\n" \
                      "$!GLOBALCONTOUR 1  LEGEND{HEADERTEXTSHAPE{ISITALIC = %s}}\n" \
                      "$!GLOBALCONTOUR 1  LEGEND{HEADERTEXTSHAPE{ISBOLD = %s}}\n" % (header_font.font_family,
                                                                                     header_font.height,
                                                                                     header_font.is_italic,
                                                                                     header_font.is_bold)
    number_settings = "$!GLOBALCONTOUR 1  LEGEND{NUMBERTEXTSHAPE{FONTFAMILY = '%s'}}\n" \
                      "$!GLOBALCONTOUR 1  LEGEND{NUMBERTEXTSHAPE{HEIGHT = %s}}\n" \
                      "$!GLOBALCONTOUR 1  LEGEND{NUMBERTEXTSHAPE{ISITALIC = %s}}\n" \
                      "$!GLOBALCONTOUR 1  LEGEND{NUMBERTEXTSHAPE{ISBOLD = %s}}\n" % (number_font.font_family,
                                                                                     number_font.height,
                                                                                     number_font.is_italic,
                                                                                     number_font.is_bold)
    result = header_settings + number_settings
    return result


def _get_axis_font_settings(x_title_font: Font = Font(), x_label_font: Font = Font(), x_title_offset=5.,
                            y_title_font: Font = Font(), y_label_font: Font = Font(), y_title_offset=5.) -> str:
    x_title = "$!TWODAXIS XDETAIL{TITLE{TEXTSHAPE{FONTFAMILY = '%s'}}}\n" \
              "$!TWODAXIS XDETAIL{TITLE{TEXTSHAPE{HEIGHT =%s}}}\n" \
              "$!TWODAXIS XDETAIL{TITLE{TEXTSHAPE{ISITALIC = %s}}}\n" \
              "$!TWODAXIS XDETAIL{TITLE{TEXTSHAPE{ISBOLD = %s}}}\n" \
              "$!TWODAXIS XDETAIL{TITLE{OFFSET = %s}}\n" % (x_title_font.font_family, x_title_font.height,
                                                            x_title_font.is_italic, x_title_font.is_bold,
                                                            x_title_offset)
    x_label = "$!TWODAXIS XDETAIL{TICKLABEL{TEXTSHAPE{FONTFAMILY = '%s'}}}\n" \
              "$!TWODAXIS XDETAIL{TICKLABEL{TEXTSHAPE{HEIGHT =%s}}}\n" \
              "$!TWODAXIS XDETAIL{TICKLABEL{TEXTSHAPE{ISITALIC = %s}}}\n" \
              "$!TWODAXIS XDETAIL{TICKLABEL{TEXTSHAPE{ISBOLD = %s}}}\n" % (x_label_font.font_family,
                                                                           x_label_font.height, x_label_font.is_italic,
                                                                           x_label_font.is_bold)
    y_title = "$!TWODAXIS YDETAIL{TITLE{TEXTSHAPE{FONTFAMILY = '%s'}}}\n" \
              "$!TWODAXIS YDETAIL{TITLE{TEXTSHAPE{HEIGHT =%s}}}\n" \
              "$!TWODAXIS YDETAIL{TITLE{TEXTSHAPE{ISITALIC = %s}}}\n" \
              "$!TWODAXIS YDETAIL{TITLE{TEXTSHAPE{ISBOLD = %s}}}\n" \
              "$!TWODAXIS YDETAIL{TITLE{OFFSET = %s}}\n" % (y_title_font.font_family, y_title_font.height,
                                                            y_title_font.is_italic, y_title_font.is_bold,
                                                            y_title_offset)
    y_label = "$!TWODAXIS YDETAIL{TICKLABEL{TEXTSHAPE{FONTFAMILY = '%s'}}}\n" \
              "$!TWODAXIS YDETAIL{TICKLABEL{TEXTSHAPE{HEIGHT =%s}}}\n" \
              "$!TWODAXIS YDETAIL{TICKLABEL{TEXTSHAPE{ISITALIC = %s}}}\n" \
              "$!TWODAXIS YDETAIL{TICKLABEL{TEXTSHAPE{ISBOLD = %s}}}\n" % (y_label_font.font_family,
                                                                           y_label_font.height, y_label_font.is_italic,
                                                                           y_label_font.is_bold)
    result = x_title + x_label + y_title + y_label
    return result


def _get_levels_setting_macro(variable_number: int, min_level, max_level,
                              num_levels: int) -> str:
    levels = np.linspace(min_level, max_level, num_levels)
    result = "$!SETCONTOURVAR\n" \
             "  VAR = %s\n" \
             "  CONTOURGROUP = 1\n" \
             "$!CONTOURLEVELS NEW\n" \
             "  CONTOURGROUP = 1\n" \
             "  RAWDATA\n" \
             "%s\n" % (variable_number, len(levels))
    template = '%s\n'
    for i in levels:
        result += template % i
    return result


def _get_legend_settings_macro(xy_position: tuple = (95, 80), rowspacing: float = 1.2, auto_levelskip: int = 1,
                               isvertical: bool = True) -> str:
    if isvertical:
        isvertical_str = 'YES'
    else:
        isvertical_str = 'NO'
    result = "$!GLOBALCONTOUR 1  LEGEND{ISVERTICAL = %s}\n" \
             "$!GLOBALCONTOUR 1  LABELS{AUTOLEVELSKIP = %s}\n" \
             "$!GLOBALCONTOUR 1  LEGEND{ROWSPACING = %s}\n" \
             "$!GLOBALCONTOUR 1\n" \
             "LEGEND\n" \
             "{\n" \
             "SHOW = YES\n" \
             "XYPOS\n" \
             "{\n" \
             "X = %s\n" \
             "Y = %s\n" \
             "}\n" \
             "}\n" % (isvertical_str, auto_levelskip, rowspacing, xy_position[0], xy_position[1])
    return result


class ColorDistribution(enum.Enum):
    BANDED = 0
    CONTINUOUS = 1


class ColorMap(enum.Enum):
    MODERN = "'Modern'"
    SMALL_RAINBOW = "'Small Rainbow'"
    WILD = "'Wild'"
    GRAY_SCALE = "GrayScale"


def _get_colormap_settings_macro(color_distribution: ColorDistribution = ColorDistribution.BANDED,
                                 colormap_name: ColorMap = ColorMap.MODERN, **kwargs) -> str:
    """
    :param color_distribution:
    :param colormap_name:
    :param kwargs: color_max и color_min, если распределение цвета =  ColorDistribution.CONTINUOUS
    :return:
    """
    string1 = "$!GLOBALCONTOUR 1  COLORMAPNAME = %s\n" \
              "$!GLOBALCONTOUR 1  COLORMAPFILTER{COLORMAPDISTRIBUTION = %s}\n" % (colormap_name.value,
                                                                                  color_distribution.name)
    if color_distribution == ColorDistribution.CONTINUOUS and 'color_min' in kwargs and 'color_max' in kwargs:
        string2 = "$!GLOBALCONTOUR 1  COLORMAPFILTER{CONTINUOUSCOLOR{CMIN = %s}}\n" \
                  "$!GLOBALCONTOUR 1  COLORMAPFILTER{CONTINUOUSCOLOR{CMAX = %s}}\n" % (kwargs['color_min'],
                                                                                       kwargs['color_max'])
    else:
        string2 = ''
    result = string1 + string2
    return result


def _get_extract_slice_command() -> str:
    return "$!CREATESLICEZONES\n"


def _get_activate_zones_command(zone_number_list: typing.List[int]) -> str:
    zones = ''
    for n, i in enumerate(zone_number_list):
        if n != len(zone_number_list) - 1:
            zones += '%s,' % i
        else:
            zones += '%s' % i
    result = "$!ACTIVEFIELDMAPS = [%s]\n" % zones
    return result


def _get_export_command(exportfname, imagewidth=1200) -> str:
    result = "$!EXPORTSETUP EXPORTFNAME = '%s'\n" \
             "$!EXPORTSETUP IMAGEWIDTH = %s\n" \
             "$!EXPORT\n" \
             "  EXPORTREGION = CURRENTFRAME\n" % (exportfname, imagewidth)
    return result


def _get_delete_zones_command(zone_number_list: typing.List[int]) -> str:
    zones = ''
    for n, i in enumerate(zone_number_list):
        if n != len(zone_number_list) - 1:
            zones += '%s,' % i
        else:
            zones += '%s' % i
    result = "$!DELETEZONES [%s]\n" % zones
    return result


def _get_show_contour_command() -> str:
    return "$!FIELDLAYERS SHOWCONTOUR = YES\n"


def _get_go_to_3d_command() -> str:
    return "$!PLOTTYPE = CARTESIAN3D\n"


def _get_frame_size_commands(width: float, height: float):
    result = "$!FRAMELAYOUT HEIGHT = %s\n" \
             "$!FRAMELAYOUT WIDTH = %s\n" % (height, width)
    return result


class FrameSettings:
    def __init__(self, width: float=9, height: float=8):
        """
        :param width:  float, optional \n
            Ширина фрейма
        :param height: float, optional \n
            Высота фрейма
        """
        self.width = width
        self.height = height


class SliceSettings:
    def __init__(self, slice_type: SliceType, position: tuple, **kwargs):
        """
        :param slice_type: SliceType \n
            определяет ориентацию секущей плоскости
        :param position: tuple \n
            кортеж из трех элементов, определяющих координаты точки, черех которую 
            проходит секущая плоскость
        :param kwargs: 1. normal - кортеж, задающий координаты нормали к секущей плоскости, необходимо задать, если
                slice_type = SliceType.ARBITRARY
        """
        self.slice_type = slice_type
        self.position = position
        self.kwargs = kwargs


class LevelSettings:
    def __init__(self, variable_number: int, min_level, max_level, num_levels: int):
        """
        :param variable_number: int \n
            номер отображаемой переменной переменной
        :param min_level: float \n
            нижняя граница отображаемого интервала значений переменной
        :param max_level: float \n
            верхняя граница отображаемого интервала значений переменной
        :param num_levels: int \
            число уровней в легенде
        """
        self.variable_number = variable_number
        self.min_level = min_level
        self.max_level = max_level
        self.num_levels = num_levels


class LegendSettings:
    def __init__(self, xy_position: tuple = (95, 80), rowspacing: float = 1.2, auto_levelskip: int = 1,
                 isvertical: bool = True, header_font: Font = Font(), number_font: Font = Font()):
        """
        :param xy_position: tuple, optional \n
            позиция легенды в координатах экрана, по умолчанию (95, 80)
        :param rowspacing: float, optional \n
            интервал между строками, по умолчанию 1.2
        :param auto_levelskip: int, optional \n
            пропуск уровней, по умолчанию 1 (без пропуска)
        :param isvertical: bool \n
            параметр, определяющей вертикальность легенды, по умолчанию True
        :param header_font: Font, optional \n
            экземпляр класса Font, содержащий настройки шрифта для заголовка легенды
        :param number_font: Font, optional \n
            экземпляр класса Font, содержащий настройки шрифта для лэйблов легенды
        """
        self.xy_position = xy_position
        self.rowspacing = rowspacing
        self.auto_levelskip = auto_levelskip
        self.isvertical = isvertical
        self.header_font = header_font
        self.number_font = number_font


class ColormapSettings:
    def __init__(self, color_distribution: ColorDistribution = ColorDistribution.BANDED,
                 colormap_name: ColorMap = ColorMap.MODERN, **kwargs):
        """
        :param color_distribution:  ColorDistribution, optional \n
            распределения цвета, по умолчанию ColorDistribution.BANDED
        :param colormap_name: ColorMap, optional \n
            цветовая схема, по умолчанию ColorMap.MODERN
        :param kwargs: color_max и color_min, если распределение цвета =  ColorDistribution.CONTINUOUS
        """
        self.color_distribution = color_distribution
        self.colormap_name = colormap_name
        self.kwargs = kwargs


class AxisSettings:
    def __init__(self, x_axis_var: int, y_axis_var: int, rect: tuple = (10, 10, 90, 90), x_line_pos: float=0,
                 y_line_pos: float=0, preserve_axis_length: bool = False,
                 x_title_font: Font = Font(), x_label_font: Font = Font(), x_title_offset: float=5.,
                 y_title_font: Font = Font(), y_label_font: Font = Font(), y_title_offset: float=5.,
                 **kwargs):
        """
        :param x_axis_var: int \n
            номер переменной, откладываемая по горизонтальной оси, например, x_axis_var = 0
        :param y_axis_var: int \n
            номер переменной, откладываемая по вертикальной оси
        :param rect: tuple, optional \n
            определяет положение прямоугольника сетки на frame, rect=(x1, y1, x2, y2),
            по умолчанию rect=(10, 10, 90, 90)
        :param preserve_axis_length: bool, optional \n
            Сохраняемость масштаба осей при изменении их диапазона
        :param x_line_pos: float \n
            Позиция горизонтальной оси по вертикали
        :param y_line_pos: float \n
            Позиция вертикальной оси по горизонтали \n
        :param x_title_font: Font, optional \n
            экземпляр класса Font, содержащий настройки шрифта для заголовка оси x
        :param x_label_font: Font, optional \n
            экземпляр класса Font, содержащий настройки шрифта для лэйблов оси x
        :param x_title_offset: float, optional \n
            сдвиг заголовка оси x относсительно оси
        :param y_title_font: Font, optional \n
            экземпляр класса Font, содержащий настройки шрифта для заголовка оси y
        :param y_label_font: Font, optional \n
            экземпляр класса Font, содержащий настройки шрифта для лэйблов оси y
        :param y_title_offset: int, optional \n
            сдвиг заголовка оси y относсительно оси
        :param kwargs: xlim и ylim (интервалы по осям x и y соотвественно), тип tuple; пример: xlim=(0,1), ylim=(1,2)
        """
        self.x_axis_var = x_axis_var
        self.y_axis_var = y_axis_var
        self.rect = rect
        self.preserve_axis_scale = preserve_axis_length
        self.x_line_pos = x_line_pos
        self.y_line_pos = y_line_pos
        self.x_title_font = x_title_font
        self.x_label_font = x_label_font
        self.x_title_offset = x_title_offset
        self.y_title_font = y_title_font
        self.y_label_font = y_label_font
        self.y_title_offset = y_title_offset
        self.kwargs = kwargs


class ExportSettings:
    def __init__(self, zone_number: int, exportfname: str, imagewidth=1200):
        """
        :param zone_number: int \n
            номер зоны, в которую будут извлечены данные среза, (на единицу большего общего
            количества зон)
        :param exportfname:  str \n
            имя файла, в который будет осуществляться экспорт
        :param imagewidth: int, optional \n
            ширина картинки
        """
        self.zone_number = zone_number
        self.exportfname = exportfname
        self.imagewidth = imagewidth


def _get_create_picture_macro(axis_settings: AxisSettings, export_settings: ExportSettings,
                              frame_settings: FrameSettings) -> str:

    extract_slice = _get_extract_slice_command()
    show_contour = _get_show_contour_command()
    go_to_2d = _get_go_to_2d_macro(axis_settings.x_axis_var, axis_settings.y_axis_var, axis_settings.x_line_pos,
                                   axis_settings.y_line_pos, axis_settings.rect,
                                   axis_settings.preserve_axis_scale, **axis_settings.kwargs)
    axis_font_settings = _get_axis_font_settings(axis_settings.x_title_font, axis_settings.x_label_font,
                                                 axis_settings.x_title_offset, axis_settings.y_title_font,
                                                 axis_settings.y_label_font, axis_settings.y_title_offset)
    activate_zone = _get_activate_zones_command([export_settings.zone_number])
    frame_size = _get_frame_size_commands(frame_settings.width, frame_settings.height)
    export = _get_export_command(export_settings.exportfname, export_settings.imagewidth)
    delete_zone = _get_delete_zones_command([export_settings.zone_number])
    go_to_3d = _get_go_to_3d_command()
    result = extract_slice + show_contour + go_to_2d + axis_font_settings + activate_zone + frame_size + export + \
             delete_zone + go_to_3d
    return result


class PictureCreator:
    def __init__(self, file_for_pictures: str, macro_filename: str, slice_settings: SliceSettings,
                 level_settings: LevelSettings, legend_settings: LegendSettings, colormap_settings: ColormapSettings,
                 axis_settings: AxisSettings, export_settings: ExportSettings, frame_settings: FrameSettings):
        """
        :param file_for_pictures: str \n
            Имя файла с расширением .plt или .lay, содержащий данные для визуализации
        :param macro_filename: tr \n
            Имя файла, в который будет сохранен макрос
        :param slice_settings: SliceSettings \n
        :param level_settings: LevelSettings \n
        :param legend_settings: LegendSettings \n
        :param colormap_settings: ColorMap \n
        :param axis_settings: AxisSettings \n
        :param export_settings: ExportSettings \n
        :param frame_settings: FrameSettings \n
        """
        self.file_for_pictures = file_for_pictures
        self.macro_filename = macro_filename
        self.slice_settings = slice_settings
        self.level_settings = level_settings
        self.legend_settings = legend_settings
        self.colormap_settings = colormap_settings
        self.axis_settings = axis_settings
        self.export_settings = export_settings
        self.frame_settings = frame_settings

    def _get_slice_settings_macro(self) -> str:
        return _get_slice_setting_macro(self.slice_settings.slice_type, self.slice_settings.position,
                                        **self.slice_settings.kwargs)

    def _get_level_settings_macro(self) -> str:
        return _get_levels_setting_macro(self.level_settings.variable_number, self.level_settings.min_level,
                                         self.level_settings.max_level, self.level_settings.num_levels)

    def _get_legend_settings_macro(self) -> str:
        return _get_legend_settings_macro(self.legend_settings.xy_position, self.legend_settings.rowspacing,
                                          self.legend_settings.auto_levelskip, self.legend_settings.isvertical)

    def _get_colormap_settings_macro(self) -> str:
        return _get_colormap_settings_macro(self.colormap_settings.color_distribution,
                                            self.colormap_settings.colormap_name, **self.colormap_settings.kwargs)

    def _get_create_picture_macro(self):
        return _get_create_picture_macro(self.axis_settings, self.export_settings, self.frame_settings)

    def _get_legend_font_settings(self):
        return _get_legend_font_settings(self.legend_settings.header_font, self.legend_settings.number_font)

    def run_creation(self):
        if os.path.splitext(self.file_for_pictures)[1] == '.plt':
            open_file = get_open_data_file_command(self.file_for_pictures)
        else:
            open_file = get_open_layout_command(self.file_for_pictures)
        slice_settings = self._get_slice_settings_macro()
        level_settings = self._get_level_settings_macro()
        legend_settings = self._get_legend_settings_macro()
        legend_font_setting = self._get_legend_font_settings()
        colormap_settings = self._get_colormap_settings_macro()
        create_picture = self._get_create_picture_macro()
        macro = wrap_macro(open_file + slice_settings + level_settings + legend_settings + legend_font_setting +
                           colormap_settings + create_picture)
        create_macro_file(macro, self.macro_filename)
        execute_macro(self.macro_filename)


if __name__ == '__main__':
    pass
    # slice_settings = SliceSettings(SliceType.ARBITRARY, (0.1, 0.15, 0.1), normal=(0, 1, 0))
    # legend_settings = LegendSettings(xy_position=(90, 45))
    # levels_settings = LevelSettings(5, 0, 90, 10)
    # colormap_settings = ColormapSettings(ColorDistribution.BANDED, ColorMap.MODERN)
    # axis_settings = AxisSettings(0, 3, rect=(10, 10, 80, 40),
    #                              xlim=(-0.1, 8.1), ylim=(-0.05, 0.4))
    # export_settings = ExportSettings(2, r'pictures\test.png')
    # pic_creator = PictureCreator(r'data_files\average_grid_density_sp_al.plt', 'macros\picture_creation.mcr',
    #                              slice_settings, levels_settings, legend_settings, colormap_settings,
    #                              axis_settings, export_settings)
    # pic_creator.run_creation()
