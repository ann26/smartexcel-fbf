import os
import unittest
from collections import namedtuple
import psycopg2
from .smart_excel import SmartExcel
# from .fbf.data_model import FbfFloodData
# from .fbf.definition import FBF_DEFINITION

# class TestFlood(unittest.TestCase):
#     def test_add_sheet(self):
#         smart_excel = SmartExcel(
#             definition=FBF_DEFINITION,
#             data=FbfFloodData(
#                 flood_event_id=15
#             )
#         )

#         smart_excel.dump()



class DataModel():
    def __init__(self):
        self.results = {
            'my_custom_payload_for_table': [
                'Good morning',
                'Bonjour'
            ],
            'my_custom_payload_for_map': [
                'Guten tag',
                'Goeie more'
            ]
        }

        self.custom_column_names = ['Bonzai', 'Artichoke']

    def __str__(self):
        return "My Custom DataModel"

    def get_sheet_name_for_summary(self):
        return 'A summary title'

    def write_column_name_func(self, instance, kwargs={}):
        return self.custom_column_names[kwargs['index']]


def get_smart_excel(definition, data_model):
    if isinstance(definition, dict):
        definition = [definition]

    return SmartExcel(
        definition=definition,
        data=data_model()
    )


class TestParseSheetDefinition(unittest.TestCase):
    def setUp(self):
        self.sheet_def = {
            'type': 'sheet',
        }

    def test_sheet_without_name(self):
        excel = get_smart_excel(self.sheet_def, DataModel)
        self.assertEqual(len(excel.sheets), 1)
        self.assertEqual(excel.sheets['Default-0-0']['name'], 'Default-0')

    def test_simple_sheet(self):
        self.sheet_def['name'] = 'Sheet 1'

        excel = get_smart_excel(self.sheet_def, DataModel)

        self.assertEqual(len(excel.sheets), 1)
        self.assertEqual(excel.sheets['Sheet 1-0']['name'], 'Sheet 1')

    def test_sheet_name_func(self):
        self.sheet_def['name'] = {
            'func': 'summary'
        }

        excel = get_smart_excel(self.sheet_def, DataModel)

        self.assertEqual(len(excel.sheets), 1)
        self.assertEqual(excel.sheets['A summary title-0']['name'], 'A summary title')

    def test_sheet_key(self):
        self.sheet_def['key'] = 'summary'
        self.sheet_def['name'] = 'A summary title'

        excel = get_smart_excel(self.sheet_def, DataModel)
        self.assertEqual(len(excel.sheets), 1)

        self.assertEqual(excel.sheets['summary']['name'], 'A summary title')

    def test_table_component(self):
        self.sheet_def['key'] = 'default'
        table_comp = {
            'type': 'table',
            'name': 'My table',
            'stack': '',
            'payload': 'my_custom_payload_for_table'
        }

        columns = [
            {
                'name': 'Column 1'
            }
        ]

        table_comp['columns'] = columns
        self.sheet_def['components'] = [
            table_comp
        ]

        excel = get_smart_excel(self.sheet_def, DataModel)

        self.assertEqual(len(excel.sheets['default']['components']), 1)
        self.assertEqual(len(excel.sheets['default']['components'][0]['columns']), 1)
        self.assertEqual(
            excel.sheets['default']['components'][0]['columns'][0],
            {
                'name': 'Column 1',
                'letter': 'A',
                'index': 0
            })

        columns = [
            {
                'name': {
                    'func': 'column_name_func'
                }
            }
        ]
        parsed_columns = excel.parse_columns(columns, repeat=1)
        self.assertEqual(
            parsed_columns[0],
            {
                'name': 'Bonzai',
                'letter': 'A',
                'index': 0
            })

    def test_map_component(self):
        self.sheet_def['key'] = 'default'
        map_comp = {
            'type': 'map',
            'name': 'My Map',
            'stack': '',
            'payload': 'my_custom_payload_for_map',
        }

        rows = [
            {
                'name': 'Row 1'
            }
        ]
        map_comp['rows'] = rows
        self.sheet_def['components'] = [
            map_comp
        ]
        excel = get_smart_excel(self.sheet_def, DataModel)
        self.assertEqual(len(excel.sheets['default']['components']), 1)
        self.assertEqual(len(excel.sheets['default']['components'][0]['rows']), 1)




class TestParseFormatDefinition(unittest.TestCase):
    def setUp(self):
        self.format_def = {
            'type': 'format',
            'key': 'my_custom_format',
            'format': {
                'border': 1,
                'bg_color': '#226b30',
            }
        }

    def test_format(self):
        excel = get_smart_excel(self.format_def, DataModel)

        self.assertEqual(len(excel.formats), 1)
        self.assertTrue('my_custom_format' in excel.formats)

    def test_num_format(self):
        self.format_def['num_format'] = 'R 0'

        excel = get_smart_excel(self.format_def, DataModel)
        self.assertEqual(len(excel.formats), 1)
        self.assertTrue('my_custom_format' in excel.formats)


class TestDump(unittest.TestCase):
    pass



if __name__ == "__main__":
    unittest.main()
