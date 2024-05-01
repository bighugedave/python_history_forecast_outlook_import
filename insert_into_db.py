import pyodbc
import sqlalchemy as db
from typing import Dict, Any, Iterable
from sqlalchemy import inspect
from datetime import datetime
from decimal import Decimal


class InsertIntoDb:
    def __init__(self):
        self.conn = None
        self.db_server = 'pulsedevsqlserver.database.windows.net'
        self.db_name = 'pulsedevdb'
        self.db_username = 'pulsedevsqlserver-admin'
        self.db_password = 'L3HSO34T18856K5O$'
        self.connectionString = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={self.db_server};DATABASE={self.db_name};UID={self.db_username};PWD={self.db_password};TrustServerCertificate=Yes'
        self.pulsedb = db.create_engine("mssql+pyodbc:///?odbc_connect={}".format(self.connectionString))

    def connect(self) -> None:
        """Estimate connection."""
        self.conn = self.pulsedb.connect()

    def insert_room_revenue(self):
        """Get list of tables."""
        inspector = inspect(self.pulsedb)
        metadata = db.MetaData()
        room_revenue = db.Table('room_revenue', metadata, autoload_with=self.pulsedb)
        # query = room_revenue.select()
        # exe_query = self.conn.execute(query)
        # query_result = exe_query.fetchmany(5)
        # print(query_result)
        query = db.insert(room_revenue).values(hotelCode='1216',
                                               room_rev_date=datetime.today(),
                                               room_rev_year=datetime.today().year,
                                               room_rev_month=datetime.today().month,
                                               day_of_week=datetime.today().weekday(),
                                               otb=0,
                                               forecast_rooms_sold=0,
                                               forecast_occupied_rooms=0,
                                               forecast_occ=Decimal('0'),
                                               forecast_room_rate=Decimal('0'),
                                               forecast_room_rev=Decimal('0'),
                                               forecast_group_rooms=0,
                                               forecast_fb_restaurants=Decimal('0'),
                                               forecast_fb_outlets=Decimal('0'),
                                               forecast_fb_banquets=Decimal('0'),
                                               actual_rooms_sold=0,
                                               available_rooms=0,
                                               actual_occ=Decimal('0'),
                                               actual_room_rate=Decimal('0'),
                                               actual_room_rev=Decimal('0'),
                                               actual_group_rooms=0,
                                               actual_fb_restaurants=Decimal('0'),
                                               actual_fb_outlets=Decimal('0'),
                                               actual_fb_banquets=Decimal('0'),
                                               budget_rooms_sold=0,
                                               budget_occ=0,
                                               budget_room_rate=0,
                                               budget_room_rev=0,
                                               budget_group_rooms=0,
                                               budget_fb_restaurants=0,
                                               budget_fb_outlets=0,
                                               budget_fb_banquets=0,
                                               report_source='',
                                               budget_fd_labor_por=Decimal('0'),
                                               nt_audit_bud_wages_par=Decimal('0'),
                                               bud_ra_dollars_por=Decimal('0'),
                                               bud_nt_hp_dollars_por=Decimal('0'),
                                               bud_ldrship_support_dollars_por=Decimal('0'),
                                               bud_leadership_support_dollars_available=Decimal('0'),
                                               total_eng_bud_dollars_par=Decimal('0'),
                                               total_sales_bud_dollars_par=Decimal('0'),
                                               total_sales_dept_sched_cost=Decimal('0'),
                                               bud_fb_dollars_por=Decimal('0'),
                                               bud_percent_fb_revenue=Decimal('0'),
                                               fb_percent_rev_dollars_avail=Decimal('0'),
                                               date_created=datetime.today(),
                                               date_modified=datetime.today()
                                               )
        query_result = self.conn.execute(query)
        self.conn.commit()

    def dispose(self) -> None:
        """Dispose opened connections."""
        self.conn.close()
        self.pulsedb.dispose()


# Testing
i = InsertIntoDb()
i.connect()
i.insert_room_revenue()
i.dispose()
