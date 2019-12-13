import os
from collections import namedtuple
import psycopg2
try:
    import plpy
except:
    pass

def namedtuplefetchall(cursor):
    "Return all rows from a cursor as a namedtuple"
    desc = cursor.description
    nt_result = namedtuple('Result', [col[0] for col in desc])
    return [nt_result(*row) for row in cursor.fetchall()]


class FbfFloodData():
    def __init__(self, flood_event_id, pl_python_env=None):
        self.flood_event_id = flood_event_id

        if pl_python_env:
            self.pl_python_env = pl_python_env
        else:
            self.connection = psycopg2.connect(
                user = os.environ['DB_USER'],
                password = os.environ['DB_PASSWORD'],
                host = os.environ['DB_HOST'],
                port = os.environ['DB_PORT'],
                database = os.environ['DB_DATABASE'])
            self.pl_python_env = False

        self.results = {
            'flood': self.get_flood(flood_event_id),
            'districts': self.get_districts(flood_event_id),
        }

    def execute_query(self, query):
        if self.pl_python_env:
            res = plpy.execute(query)
            try:
                fields = list(res[0].keys())

                def_meta_res = namedtuple('Result', ', '.join(fields))

                results = [
                    def_meta_res(*list(res[index].values()))
                    for index in range(0, len(res))
                ]
            except IndexError:
                results = []
            return results

        else:
            with self.connection.cursor() as cursor:
                cursor.execute(query)
                results = namedtuplefetchall(cursor)

        return results

    def get_districts(self, flood_event_id):

        query = """
            SELECT
                area.name as district_name,
                area.dc_code as district_code,
                summary.vulnerability_total_score as vulnerability_total_score,
                summary.building_count as total_buildings,
                summary.flooded_building_count as flooded_buildings,
                summary.residential_building_count as residential_building_count,
                summary.residential_flooded_building_count as residential_flooded_building_count,
                summary.clinic_dr_building_count as clinic_dr_building_count,
                summary.clinic_dr_flooded_building_count as clinic_dr_flooded_building_count

            FROM
                flood_event fe,
                flood_event_district_summary summary,
                district area
            WHERE
                fe.id = summary.flood_event_id
                and summary.district_id = area.dc_code
                and fe.id = {flood_event_id}
            ;
        """.format(
            flood_event_id=flood_event_id
        )

        return self.execute_query(query)


    def get_subdistricts(self, flood_event_id, district_code):

        query = """
            SELECT
                area.name as sub_district_name,
                area.sub_dc_code as sub_district_code,
                summary.vulnerability_total_score as vulnerability_total_score,
                summary.building_count as total_buildings,
                summary.flooded_building_count as flooded_buildings,
                summary.residential_building_count as residential_building_count,
                summary.residential_flooded_building_count as residential_flooded_building_count,
                summary.clinic_dr_building_count as clinic_dr_building_count,
                summary.clinic_dr_flooded_building_count as clinic_dr_flooded_building_count

            FROM
                flood_event fe,
                flood_event_sub_district_summary summary,
                sub_district area
            WHERE
                fe.id = summary.flood_event_id
                and summary.sub_district_id = area.sub_dc_code
                and area.dc_code = {district_code}
                and fe.id = {flood_event_id}
            ;
        """.format(
            district_code=district_code,
            flood_event_id=flood_event_id
        )

        return self.execute_query(query)


    def get_villages(self, flood_event_id, sub_district_code):
        query = """
            SELECT
                area.name as village_name,
                area.village_code as village_code,
                summary.vulnerability_total_score as vulnerability_total_score,
                summary.building_count as total_buildings,
                summary.flooded_building_count as flooded_buildings,
                summary.residential_building_count as residential_building_count,
                summary.residential_flooded_building_count as residential_flooded_building_count,
                summary.clinic_dr_building_count as clinic_dr_building_count,
                summary.clinic_dr_flooded_building_count as clinic_dr_flooded_building_count

            FROM
                flood_event fe,
                flood_event_village_summary summary,
                village area
            WHERE
                fe.id = summary.flood_event_id
                and summary.village_id = area.village_code
                and area.sub_dc_code = {sub_district_code}
                and fe.id = {flood_event_id}
            ;
        """.format(
            sub_district_code=sub_district_code,
            flood_event_id=flood_event_id
        )

        return self.execute_query(query)


    def get_flood(self, flood_event_id):
        query = """
            SELECT
                *
            FROM
                flood_event fe
            WHERE
                fe.id = {flood_event_id}
        """.format(
            flood_event_id=flood_event_id
        )

        return self.execute_query(query)

    def get_payload_subdistricts(self, instance, foreign_key):
        district_code = int(getattr(instance, foreign_key))
        return self.get_subdistricts(self.flood_event_id, district_code)

    def get_payload_villages(self, instance, foreign_key):
        sub_district_code = int(getattr(instance, foreign_key))
        return self.get_villages(self.flood_event_id, sub_district_code)

    def get_sheet_name_for_flood_summary(self, kwargs={}):
        return f"Flood Summary"

    def get_sheet_name_for_subdistrict_summary(self, instance, kwargs={}):
        name = f"District {instance.district_name} Summary"
        return name[0:30]

    def get_sheet_name_for_village_summary(self, instance, kwargs={}):
        name = f"Sub district {instance.sub_district_name} Summary"
        return name[0:30]

    def write_flood_title(self, instance, kwargs={}):
        return f"Flood {instance.id}"

    def write_flood_acquisition_date(self, instance, kwargs={}):
        try:
            if isinstance(instance.acquisition_date, str):
                return instance.acquisition_date
            else:
                return instance.acquisition_date.strftime('%Y-%m-%d')
        except Exception:
            return None

    def write_flood_forecast_date(self, instance, kwargs={}):
        try:
            if isinstance(instance.forecast_date, str):
                return instance.forecast_date
            else:
                return instance.forecast_date.strftime('%Y-%m-%d')
        except Exception:
            return None

    def write_flood_source(self, instance, kwargs={}):
        return instance.source

    def write_flood_notes(self, instance, kwargs={}):
        return instance.notes

    def write_flood_link(self, instance, kwargs={}):
        return instance.link

    def write_flood_trigger_status(self, instance, kwargs={}):
        return instance.trigger_status

    def write_district_name(self, instance, kwargs={}):
        return instance.district_name

    def write_district_code(self, instance, kwargs={}):
        return instance.district_code

    def write_sub_district_name(self, instance, kwargs={}):
        return instance.sub_district_name

    def write_sub_district_id(self, instance, kwargs={}):
        return instance.sub_district_code

    def write_village_name(self, instance, kwargs={}):
        return instance.village_name

    def write_village_id(self, instance, kwargs={}):
        # village_id : 3201160018.0
        # cast to int to remove the decimal
        # then to str because 3201160018 is bigger than 2,147,483,647
        return str(int(instance.village_code))

    def write_total_buildings(self, instance, kwargs={}):
        return instance.total_buildings

    def write_flooded_buildings(self, instance, kwargs={}):
        return instance.flooded_buildings

    def write_not_flooded_buildings(self, instance, kwargs={}):
        try:
            return instance.total_buildings - instance.flooded_buildings
        except Exception:
            return 0

    def write_vulnerability_total_score(self, instance, kwargs={}):
        return instance[kwargs['index']]['vulnerability_total_score']

    def write_building_count(self, instance, kwargs={}):
        return instance.building_count

    def write_flooded_building_count(self, instance, kwargs={}):
        return instance.flooded_building_count

    def write_residential_building_count(self, instance, kwargs={}):
        return instance.residential_building_count

    def write_residential_flooded_building_count(self, instance, kwargs={}):
        return instance.residential_flooded_building_count

    def write_clinic_dr_building_count(self, instance, kwargs={}):
        return instance.clinic_dr_building_count

    def write_clinic_dr_flooded_building_count(self, instance, kwargs={}):
        return instance.clinic_dr_flooded_building_count

