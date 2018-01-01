from distutils.core import setup

# Long description to be published in PyPi
LONG_DESCRIPTION = """
**pyppt**: Python interface for adding figures to Microsoft PowerPoint presentations on-the-fly.

For the documentation please refer to README.md inside the package or on the
GitHub (https://github.com/vfilimonov/pyppt/blob/master/README.md).
"""

_URL = 'http://github.com/vfilimonov/pyppt'
_VERSION = '0.1'

setup(name='pyppt',
      version=_VERSION,
      description='Python interface for adding figures to Microsoft PowerPoint presentations on-the-fly.',
      long_description=LONG_DESCRIPTION,
      url=_URL,
      download_url=_URL + '/archive/v' + _VERSION + '.zip',
      author='Vladimir Filimonov',
      author_email='vladimir.a.filimonov@gmail.com',
      license='MIT License',
      packages=['pyppt'],
      install_requires=[],
      classifiers=['Programming Language :: Python :: 2',
                   'Programming Language :: Python :: 3', ]
      )
