# SmartExcel for Fbf

Report generator for the Fbf project.

## Documentation
Read the code and the tests.

## Development
```
python -m venv env
source env/bin/activate
pip install -r requirements.txt
```

### environment variables
```
cp .env.template .env
```

Edit `.env` then
```
source .env
```

## Test
There aren't proper tests...
It is a simple way to run code.
`TestFlood` needs access to the Fbf database.

```
python -m smartexcel.test_smart_excel TestFlood
```
A file `template.xlsx` is created at the root of this project. Open it and make sure it looks like what you want.


## Deployment


### SmartExcel code
1. Do your change, edit the definition, tweak the data model. Make sure it works by running the test. Then commit and push.
2. Go to rancher > Fbf stack > connect to the shell of the `db` container
3. Run these commands:
```
cd /usr/local/lib/python3.7/dist-packages
rm -rf SmartExcel
git clone https://github.com/pierrealixt/SmartExcel.git
```

### python/pl code

The function's name is `fbf_generate_excel_report_for_flood`. It is already present in the database.
If you need to change the code of the function, you must replace it.

```
CREATE OR REPLACE FUNCTION fbf_generate_excel_report_for_flood (flood_event_id integer)
  RETURNS varchar
 AS $$
   import io
   import sys
   sys.path.insert(0, '/usr/local/lib/python3.7/dist-packages')
   sys.path.insert(0, '/usr/local/lib/python3.7/dist-packages/SmartExcel/smartexcel')
   from smart_excel import SmartExcel
   from fbf.data_model import FbfFloodData
   from fbf.definition import FBF_DEFINITION

   excel = SmartExcel(
       output=io.BytesIO(),
       definition=FBF_DEFINITION,
       data=FbfFloodData(
           flood_event_id=flood_event_id,
           pl_python_env=True
       )
   )
   excel.dump()

   plan = plpy.prepare("UPDATE flood_event SET spreadsheet = ($1) where id = ($2)", ["bytea", "integer"])
   plpy.execute(plan, [excel.output.getvalue(), flood_event_id])

   return "OK"
$$ LANGUAGE plpython3u;
```


## How to use it

```
select * from fbf_generate_excel_report_for_flood(15);
```
`15` is an ID present in the table `flood_event`.
