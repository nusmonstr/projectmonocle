except ImportError:
    from distutils.core import setup

config = {
    'description': 'My Project',
    'author': 'Eric Champe',
    'url': 'https://github.com/nusmonstr/projectmonocle',
    'author_email': 'eric_champe@live.com',
    'version': '1.1',
    'install_requires': ['win32com'],
    'packages': [''],
    'scripts': [''],
    'name': 'FinPy - Personal Finance Aggregator'
}

setup(**config)