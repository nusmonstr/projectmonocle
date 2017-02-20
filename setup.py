except ImportError:
    from distutils.core import setup

config = {
    'description': 'My Project',
    'author': 'Eric Champe',
    'url': 'url to get it at',
    'download_url': 'where to download it',
    'author_email': 'eric_champe@live.com',
    'version': '0.1',
    'install_requires': ['nose'],
    'packages': ['NAME'],
    'scripts': [],
    'name': 'projectname'
}

setup(**config)