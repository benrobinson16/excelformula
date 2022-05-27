import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name='excelformula',
    version='0.0.1',
    author='Ben Robinson',
    author_email='hello@benrobinson.dev',
    description='Database Integration of Excel Formulae',
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/benrobinson16/excelformula',
    packages=['excelformula'],
    install_requires=[
        'formulas @ git+https://github.com/benrobinson16/formulas.git@master',
        'openpyxl',
        'numpy',
        'json'
    ],
)