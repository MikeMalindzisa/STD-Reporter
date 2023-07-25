from PyInstaller.utils.hooks import collect_data_files, copy_metadata

# Include the icon file and settings.py as data files
datas = collect_data_files('.', include_py_files=True)

# Set the version, author, and publisher metadata
hiddenimports = []
copy_metadata(hiddenimports)

# Set the metadata explicitly
metadata = {
    'version': '1.0.13',
    'author': 'Mike Malindzisa',
    'publisher': 'Standard Bank',
    'name': 'STD Reporter',
    'icon': 'icon.png',
}
