import os
from setuptools import setup, find_packages

def read(fname):
    return open(os.path.join(os.path.dirname(__file__), fname)).read()

setup(
    name = "transcription",
    version = "1.3",
    author = "Felipe Valverde",
    author_email = "felipe.valverde.campos@gmail.com",
    keywords = "transcription sheet tk",
    url = "https://github.com/fvalverd/Transcription.git",
    packages = find_packages(exclude=['ez_setup', 'examples', 'tests']),
    scripts = ['scripts/transcription_executable.py'],
    include_package_data = True,
    zip_safe = False,
    long_description = read('README'),
    install_requires = [
          "openpyxl",
      ],
)