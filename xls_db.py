from typing import Optional
from sqlalchemy import create_engine, select, desc
from sqlalchemy.orm import DeclarativeBase, Session, Mapped, mapped_column
from sqlalchemy.dialects.sqlite import DATE
import datetime

""" Module for connect to DB """


class Base(DeclarativeBase):
    pass


class XlsStorage:
    """ Database class """
    class XlsTable(Base):
        __tablename__ = "xls_data"
        id: Mapped[int] = mapped_column(primary_key=True)
        xls_id: Mapped[int] = mapped_column()
        company: Mapped[Optional[str]]
        fact_qliq: Mapped[int] = mapped_column()
        fact_qoil: Mapped[int] = mapped_column()
        forecast_qliq: Mapped[int] = mapped_column()
        forecast_qoil: Mapped[int] = mapped_column()
        date: Mapped[DATE] = mapped_column(DATE)
        total_qliq: Mapped[int] = mapped_column()
        total_qoil: Mapped[int] = mapped_column()

        def __repr__(self) -> str:
            return f"xls_data(id={self.id!r}, xls_id={self.xls_id!r}, company={self.company!r}, " \
                   f"fact_qliq={self.fact_qliq!r}, fact_qoil={self.fact_qoil!r}, data={self.date!r}, " \
                   f"total_qliq={self.total_qliq!r}, total_qoil={self.total_qoil!r})"

    def __init__(self, path):
        self.database_engine = create_engine(f'sqlite:///{path}', echo=False,)
        self.metadata = Base.metadata.create_all(self.database_engine)
        self.XlsTable()
        self.session = Session(self.database_engine)
        self.session.commit()

    def add_row(self, xls_id, company, fact_qliq, fact_qoil, forecast_qliq,
                forecast_qoil, date, total_qliq, total_qoil):
        self.session = Session(self.database_engine)
        row = self.XlsTable(xls_id=xls_id,
                            company=company,
                            fact_qliq=fact_qliq,
                            fact_qoil=fact_qoil,
                            forecast_qliq=forecast_qliq,
                            forecast_qoil=forecast_qoil,
                            date=date,
                            total_qliq=total_qliq,
                            total_qoil=total_qoil)
        self.session.add_all([row])
        self.session.commit()


if __name__ == '__main__':
    test_db = XlsStorage('test_xls_db.db3')
    date_test = datetime.date(2023, 3, 30)
    date = datetime.datetime.now().date() + datetime.timedelta(days=-1)
    date_now = datetime.datetime.now().date()
    test_db.add_row(1, 'company2', 10, 30, 40, 22, date_test, 27, 28)
    test_db.add_row(1, 'company1', 10, 30, 40, 22, date_test, 27, 28)
    test_db.add_row(1, 'company2', 10, 30, 40, 22, date, 27, 28)
    rows_list = select(test_db.XlsTable).order_by(desc(test_db.XlsTable.id)).limit(1)
    rows_list_2 = select(test_db.XlsTable).where(
                                            test_db.XlsTable.date.is_(date),
                                            test_db.XlsTable.company.is_('company2')
                                            ).order_by(desc(test_db.XlsTable.id)).limit(1)
    rows = test_db.session.scalars(rows_list)
    rows_2 = test_db.session.scalars(rows_list_2)
    # for row in rows:
    #     print(row, row.company)
    for row in rows_2:
        print(row, row.company)
