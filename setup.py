from setuptools import setup


author = 'Zncl2222'
email = 'zwebapplication@gmail.com'
url = 'https://github.com/Zncl2222/openpyxl_style_writer'

with open('README.md', 'r', encoding='utf-8') as fh:
    long_description = fh.read()

setup(
    name='openpyxl_style_writer',
    version='1.1.0',
    description='A wrapper for openpyxl to create and use resualbe style in write only mode',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url=url,
    author=author,
    author_email=email,
    license='MIT',
    install_requires=['openpyxl'],
    keywords=['openpyxl', 'excel', 'style'],
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
    ],
    python_requires='>=3.6',
)
