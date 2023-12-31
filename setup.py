from setuptools import setup, find_packages

with open('LONGDESCRIPTION.rst', 'r', encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='FastXlsToCsv',
    version='1.2.1',
    description='A fast way to convert .xls and .xlsx to csv with vbs',
    long_description= long_description,
    url= 'https://github.com/willayy/FastXlsToCsv',
    license= 'MIT',
    author_email= 'williamnorland@gmail.com',
    author='William Norland',
    packages=find_packages(),
    install_requires=[],  # Add any dependencies here,
    classifiers= ['Operating System :: Microsoft :: Windows',
                  'Operating System :: Microsoft :: Windows :: Windows 10',
                  'Operating System :: Microsoft :: Windows :: Windows 11',
                  'License :: OSI Approved :: MIT License',
                  'Programming Language :: Python',
                  'Programming Language :: Python :: 3',
                  'Programming Language :: Python :: 3.11']
)