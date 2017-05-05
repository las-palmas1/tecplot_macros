from distutils.core import setup
from tecplot_lib import __author__, __version__

setup(
    name='tecplot_lib',
    version=__version__,
    py_modules=['tecplot_lib'],
    url='',
    license='',
    author=__author__,
    author_email='',
    description='Python wrap for some TecPlot scripting language functions'
)

