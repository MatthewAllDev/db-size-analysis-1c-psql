from db_size_analysis_1c_psql import create_db_size_analysis
import getpass

if __name__ == '__main__':
    host: str = '192.168.1.100'
    database: str = input('DataBase name: ').strip()
    user: str = input('User name: ').strip()
    password: str = getpass.getpass('Password: ').strip()
    analysis = create_db_size_analysis(host, database, user, password, '1C_struct.json')
    analysis.sort('size', True)
    analysis.export_to_excel('export.xlsx')
