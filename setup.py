from distutils.core import setup

import sys
sys.path.insert(0, './pyppt')
from _ver_ import __version__, __author__, __email__, __url__

# Long description to be published in PyPi
LONG_DESCRIPTION = """
**pyppt**: Python interface for adding figures to Microsoft PowerPoint presentations on-the-fly.

For the documentation please refer to README.md inside the package or on the
GitHub (%s/blob/master/README.md).
""" % (__url__)

setup(name='pyppt',
      version=__version__,
      description='Python interface for adding figures to Microsoft PowerPoint presentations on-the-fly',
      long_description=LONG_DESCRIPTION,
      url=__url__,
      download_url=__url__ + '/archive/v' + __version__ + '.zip',
      author=__author__,
      author_email=__email__,
      license='MIT License',
      packages=['pyppt'],
      install_requires=['future'],
      include_package_data=True,
      entry_points={
          'console_scripts': ['pyppt_server = pyppt.server:pyppt_server'],
      },
      classifiers=['Programming Language :: Python :: 2',
                   'Programming Language :: Python :: 3', ]
      )
