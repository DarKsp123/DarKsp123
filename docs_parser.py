from spire.doc import Document, FileFormat, Table, SpireException
import shutil
from spire.doc.common import *


file_path = r'P:\99999_2400_RESEARCH\1. Research Nikita\Разное\Договор найма квартиры Чака\2023_Договор найма_ПРОЕКТ_ред_рецензирование 27.02.doc'
# file_path = r'P:\50016_1001_Новая форма заявления в подкомиссию_ИХ3_15112022_v2.docx'

# suffix = ''
# if file_path.split('.')[-1] == 'doc':
#     suffix = FileFormat.Doc
# else:
#     suffix = FileFormat.Docx

file_name = file_path.split('\\')[-1]
print("Документ: ", file_name)
temp_file_path = f'temp_{file_name}'

if os.path.exists(file_path):
    try:
        # Создаем временную копию файла
        shutil.copyfile(file_path, temp_file_path)

        # Загружаем временную копию файла
        document = Document()
        document.LoadFromFile(temp_file_path)
        all_text = ''
        print(f'Кол-во разделов: {document.Sections.Count}')
        # Обработка документа
        for s in range(document.Sections.Count):
            # Sections: кол-во разделов в документе
            section = document.Sections.get_Item(s)
            all_text += '\n'.join([section.Paragraphs.get_Item(i).Text for i in range(section.Paragraphs.Count)])
            tables = section.Tables  # таблицы в разделе документа
            print('Кол-во таблиц в разделе: ', tables.Count)
            for i in range(0, tables.Count):
                table = tables.get_Item(i)  # забираем объект "Таблица"
                tableData = ''
                print(f'Кол-во строк в таблице: {table.Rows.Count}')
                for j in range(0, table.Rows.Count):
                    print(f'Кол-во столбцов в строке {j}: {table.Rows.get_Item(j).Cells.Count}')
                    # Loop through the cells of the row
                    for k in range(0, table.Rows.get_Item(j).Cells.Count):
                        # Get a cell
                        cell = table.Rows.get_Item(j).Cells.get_Item(k)
                        # Get the text in the cell
                        cellText = ''
                        for para in range(cell.Paragraphs.Count):
                            paragraphText = cell.Paragraphs.get_Item(para).Text
                            # print(paragraphText)
                            cellText += paragraphText + ' '
                        # Add the text to the string
                        tableData += cellText
                        if k <= table.Rows.get_Item(j).Cells.Count - 1:
                            # tableData += '\t'
                            # Add a new line
                            tableData += '\n'
                        # print(cellText)
                # print(tableData)
                all_text += tableData
        print(all_text)

    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'document' in locals() and document is not None:
            document.Close()
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
