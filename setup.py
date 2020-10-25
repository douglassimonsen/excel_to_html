import setuptools


setuptools.setup(
    name="excel_to_html",
    packages=['excel_to_html'],
    version="1.0.0",
    description="A package that converts excel sheets to HTML tables",
    url="https://github.com/mwhamilton/excel_to_html",
    download_url="https://github.com/mwhamilton/excel_to_html/archive/1.0.0.tar.gz",
    author="Matthew Hamilton",
    author_email="mwhamilton6@gmail.com",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
    ],
    include_package_data=True,
    install_requires=[
        "colorsys",
        "openpyxl",
        "jinja2",
    ],
)
