import logic.export_excel

ex = logic.export_excel.ExportExcel('test.json')
ex.import_from_files()
ex.exportAsExcel()
