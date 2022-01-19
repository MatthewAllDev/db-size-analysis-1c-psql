import psycopg2
from contextlib import closing
import json
from .data_base_object import DataBaseObject
from .safe_list import SafeList


def get_psql_objects_with_size(host: str, database: str, user: str, password: str) -> SafeList:
    with closing(psycopg2.connect(host=host, dbname=database, user=user, password=password)) as conn:
        request: str = "SELECT nspname || '.' || relname AS \"relation\", " \
                       "pg_total_relation_size(C.oid) AS \"total_size\" " \
                       "FROM pg_class C " \
                       "LEFT JOIN pg_namespace N ON (N.oid = C.relnamespace) " \
                       "WHERE nspname NOT IN ('pg_catalog', 'information_schema') " \
                       "AND C.relkind <> 'i' " \
                       "AND nspname !~ '^pg_toast' " \
                       "ORDER BY nspname || '.' || relname;"
        db_objects: SafeList = SafeList()
        with conn.cursor() as cursor:
            cursor.execute(request)
            for row in cursor:
                row: tuple[str, int]
                names: SafeList[str] = SafeList(row[0].replace('_', '._').split('.'))
                db_objects.append(DataBaseObject(namespace=names.get(0),
                                                 parent_object=names.get(2),
                                                 object=names.get(2, names.get(1)) + names.get(3, ''),
                                                 size=row[1]))
    return db_objects


def load_1c_database_struct(file_path: str) -> dict:
    with open(file_path, 'r', encoding='utf-8') as file:
        db_struct: dict = json.load(file)
    return db_struct


def compile_db_objects_with_1c_struct(db_objects: SafeList, struct_1c: dict, create_tree: bool = False) -> \
        SafeList or DataBaseObject:
    data: DataBaseObject = DataBaseObject('DataBase')
    for db_object in db_objects:
        db_object.set_attributes(**struct_1c.get(db_object.object, dict()))
        if create_tree:
            if hasattr(db_object, 'table_name_1c') and db_object.table_name_1c != '':
                keys: SafeList = SafeList(db_object.table_name_1c.split('.'))
            elif hasattr(db_object, 'metadata') and db_object.metadata != '':
                keys: SafeList = SafeList(db_object.metadata.split('.'))
                if hasattr(db_object, 'purpose') and db_object.purpose != '':
                    keys.append(db_object.purpose)
            else:
                keys: SafeList = SafeList()
                keys.append('Service')
                keys.append(db_object.object)
            data.add_data_level(keys, db_object)
    if create_tree:
        return data
    else:
        return db_objects


def create_db_size_analysis(host: str, database: str, user: str, password: str, struct_1c_file_path: str):
    db_objects: SafeList = get_psql_objects_with_size(host, database, user, password)
    struct_1c: dict = load_1c_database_struct(struct_1c_file_path)
    return compile_db_objects_with_1c_struct(db_objects, struct_1c, True)
