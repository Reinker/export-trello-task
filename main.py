import exportExcel.export_excel

ex = exportExcel.export_excel.ExportExcel('test.json')
ex.import_from_files()
ex.exportAsExcel()
