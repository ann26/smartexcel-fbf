import os
import unittest
from .smart_excel import (
    SmartExcel,
    validate_position
)


children_things = [
    {
        'parent_id': 42,
        'result': 'yes'
    },
    {
        'parent_id': 42,
        'result': 'no'
    },
    {
        'parent_id': 43,
        'result': 'oui'
    },
    {
        'parent_id': 43,
        'result': 'non'
    },
]

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
            ],
            'things': [
                {
                    'id': 42,
                    'name': 'The answer'
                },
                {
                    'id': 43,
                    'name': 'nothing'
                }
            ]
        }

        self.custom_column_names = ['Bonzai', 'Artichoke']

    def __str__(self):
        return "My Custom DataModel"

    def get_sheet_name_for_summary(self):
        return 'A summary title'

    def write_column_name_func(self, instance, kwargs={}):
        return self.custom_column_names[kwargs['index']]

    def write_first_column(self, instance, kwargs={}):
        return instance

    def get_payload_detail(self, instance, foreign_key):
        item_id = instance[foreign_key]

        return [
            item
            for item in children_things
            if item['parent_id'] == item_id
        ]

    def get_sheet_name_for_detail(self, instance):
        return f'Sheet nb {instance["id"]}'

    def write_thing_id(self, instance, kwargs={}):
        return instance['id']

    def write_thing_value(self, instance, kwargs={}):
        return instance['name']

    def write_result(self, instance, kwargs={}):
        return instance['result']


def get_smart_excel(definition, data_model, output='template.xlsx'):
    if isinstance(definition, dict):
        definition = [definition]

    return SmartExcel(
        output=output,
        definition=definition,
        data=data_model()
    )


class TestParseSheetDefinition(unittest.TestCase):
    def setUp(self):
        self.sheet_def = {
            'type': 'sheet',
        }

    def test_reserved_sheet(self):
        excel = get_smart_excel({}, DataModel)
        self.assertEqual(len(excel.sheets), 2)
        self.assertTrue('_data' in list(excel.sheets.keys()))
        self.assertTrue('_meta' in list(excel.sheets.keys()))

        self.assertTrue(excel.sheets['_data']['reserved'])
        self.assertTrue(excel.sheets['_meta']['reserved'])

        # user should not be able to add a sheet
        # with a reserved name (_data, _meta)
        for reserved_sheet_name in excel.reserved_sheets:
            self.sheet_def['name'] = reserved_sheet_name

            with self.assertRaises(ValueError) as raised:
                excel = get_smart_excel(self.sheet_def, DataModel)

            self.assertEqual(str(raised.exception), f'{reserved_sheet_name} is a reserved sheet name.')
            self.assertEqual(len(excel.sheets), 2)


    def test_sheet_without_name(self):
        excel = get_smart_excel(self.sheet_def, DataModel)
        self.assertEqual(len(excel.sheets), 3)
        self.assertEqual(excel.sheets['Default-0-0']['name'], 'Default-0')
        self.assertFalse(excel.sheets['Default-0-0']['reserved'])

    def test_simple_sheet(self):
        self.sheet_def['name'] = 'Sheet 1'

        excel = get_smart_excel(self.sheet_def, DataModel)

        self.assertEqual(len(excel.sheets), 3)
        self.assertEqual(excel.sheets['Sheet 1-0']['name'], 'Sheet 1')

    def test_sheet_name_func(self):
        self.sheet_def['name'] = {
            'func': 'summary'
        }

        excel = get_smart_excel(self.sheet_def, DataModel)

        self.assertEqual(len(excel.sheets), 3)
        self.assertEqual(excel.sheets['A summary title-0']['name'], 'A summary title')

    def test_sheet_key(self):
        self.sheet_def['key'] = 'summary'
        self.sheet_def['name'] = 'A summary title'

        excel = get_smart_excel(self.sheet_def, DataModel)
        self.assertEqual(len(excel.sheets), 3)

        self.assertEqual(excel.sheets['summary']['name'], 'A summary title')

    def test_table_component(self):
        self.sheet_def['key'] = 'default'
        table_comp = {
            'type': 'table',
            'name': 'My table',
            'position': {
                'x': 0,
                'y': 0
            },
            'payload': 'my_custom_payload_for_table'
        }

        columns = [
            {
                'name': 'Column 1',
                'key': 'column_1'
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
                'index': 0,
                'key': 'column_1'
            })

        columns = [
            {
                'name': {
                    'func': 'column_name_func'
                },
                'key': 'column_1'
            }
        ]
        parsed_columns = excel.parse_columns(columns, repeat=1)
        self.assertEqual(
            parsed_columns[0],
            {
                'name': 'Bonzai',
                'letter': 'A',
                'index': 0,
                'key': 'column_1'
            })

    def test_map_component(self):
        self.sheet_def['key'] = 'default'
        map_comp = {
            'type': 'map',
            'name': 'My Map',
            'position': {
                'x': 0,
                'y': 0
            },
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

    def test_component_position(self):
        self.sheet_def['key'] = 'default'
        components = [
            {
                'type': 'map',
                'name': 'My Map',
                'position': {
                    'x': 0,
                    'y': 0
                },
                'payload': 'my_custom_payload_for_map',
                'rows': [
                    {
                        'name': 'Row 1'
                    }
                ]
            },
            {
                'type': 'map',
                'name': 'a second map',
                'position': {
                    'x': 0,
                    'y': 0
                }
            }
        ]

        self.sheet_def['components'] = components
        with self.assertRaises(ValueError) as raised:
            excel = get_smart_excel(self.sheet_def, DataModel)

        self.assertEqual(str(raised.exception), 'Cannot position `a second map` at 0;0. `My Map` is already present.')

        self.sheet_def['components'][1]['position']['y'] = 1
        excel = get_smart_excel(self.sheet_def, DataModel)

    def test_recursive(self):
        self.sheet_def['components'] = [
            {
                'type': 'table',
                'name': 'A table',
                'payload': 'things',
                'position': {
                    'x': 0,
                    'y':0
                },
                'columns': [
                    {
                        'name': 'Identification',
                        'key': 'thing_id'
                    },
                    {
                        'name': 'Value',
                        'key': 'thing_value'
                    }
                ],
                'recursive': {
                    'payload_func': 'detail',
                    'foreign_key': 'id',
                    'name': {
                        'func': 'detail'
                    },
                    'components': [
                        {
                            'name': 'Another table',
                            'type': 'table',
                            'position': {
                                'x': 0,
                                'y': 0
                            },
                            'columns': [
                                {
                                    'name': 'Result',
                                    'key': 'result'
                                }
                            ]
                        }
                    ]
                }
            }
        ]
        excel = get_smart_excel(
            self.sheet_def,
            DataModel,
            'test_recursive.xlsx')

        self.assertEqual(len(excel.sheets), 5)
        excel.dump()

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
    def runTest(self):
        path = os.path.join(
            os.path.dirname(os.path.realpath(__file__)),
            'test_dump.xlsx')

        if os.path.exists(path):
            os.remove(path)

        self.format_def = {
            'type': 'sheet',
            'name': 'Bonjour',
            'components': [
                {
                    'type': 'table',
                    'name': 'My table',
                    'position': {
                        'x': 0,
                        'y': 0
                    },
                    'payload': 'my_custom_payload_for_table',
                    'columns': [
                        {
                            'name': 'Column 1',
                            'key': 'first_column'
                        }
                    ]
                }
            ]
        }


        excel = get_smart_excel(
            self.format_def,
            DataModel,
            output=path)
        excel.dump()

        self.assertTrue(os.path.exists(path))

        # Do a manuel testing of the generated spreadsheet.


class TestValidatePosition(unittest.TestCase):
    def setUp(self):
        self.element = {
            'position': ''
        }

    def test_type(self):
        with self.assertRaises(ValueError) as raised:
            validate_position(self.element)

        self.assertEqual(str(raised.exception), "position must be a <class 'dict'>")

    def test_required_attrs(self):
        self.element['position'] = {}

        with self.assertRaises(ValueError) as raised:
            validate_position(self.element)

        self.assertEqual(str(raised.exception), "x is required in a component position definition.")

        self.element['position'] = {
            'x': None
        }

        with self.assertRaises(ValueError) as raised:
            validate_position(self.element)

        self.assertEqual(str(raised.exception), "y is required in a component position definition.")

    def test_type_required_attrs(self):
        self.element['position'] = {
            'x': None,
            'y': None
        }

        with self.assertRaises(ValueError) as raised:
            validate_position(self.element)

        self.assertEqual(str(raised.exception), "x must be a <class 'int'>")

        self.element['position'] = {
            'x': 0,
            'y': None
        }

        with self.assertRaises(ValueError) as raised:
            validate_position(self.element)

        self.assertEqual(str(raised.exception), "y must be a <class 'int'>")

    def test_ok(self):
        self.element['position'] = {
            'x': 0,
            'y': 0
        }

        self.assertTrue(validate_position(self.element))


if __name__ == "__main__":
    unittest.main()
