try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = {
    'description': 'Project name',
    'author': 'author name',
    'url': 'Where to get it',
    'download_url': 'Where to download it',
    'author_email': 'author email',
    'version': '1.0',
    'install_requires': ['nose'],
    'pakages': ['NAME'],
    'scripts': [],
    'name': 'project name'
}

setup(**config)
