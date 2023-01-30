import pandas as pd
from datetime import datetime

from sqlalchemy import create_engine, Column, Date, ForeignKey, Integer, String, text, func
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import declared_attr, Session, relationship
from sqlalchemy.future import select

class Base:
    @declared_attr
    def __tablename__(cls):
        return cls.__name__.lower()

    __table_args__ = {"mysql_engine": "InnoDB"}

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)


Base = declarative_base(cls=Base)

engine = create_engine('mysql+pymysql://hkjc:Fr5cypE3D%@mysql.ixx.cc:3306/hkjc')



class Marksix(Base):
    __tablename__ = "MarkSix"

    DrawNo = Column(String(6))
    Date = Column(Date)
    DrawnNo1 = Column(Integer)
    DrawnNo2 = Column(Integer)
    DrawnNo3 = Column(Integer)
    DrawnNo4 = Column(Integer)
    DrawnNo5 = Column(Integer)
    DrawnNo6 = Column(Integer)
    ExtraNum = Column(Integer)

with Session(engine) as session:
    with open('2023.txt') as f:
        lines = f.readlines()
        for line in lines:
            # print(line)
            cols = line.split("\t")

            marksix = Marksix()
            # print(cols)
            if len(cols)==9:
                marksix.DrawNo = cols[0][2:4]+'-'+cols[0][4:7]
                marksix.Date = datetime.strptime(cols[1], "%Y/%m/%d")

                marksix.DrawnNo1 = int(cols[2])
                marksix.DrawnNo2= int(cols[3])
                marksix.DrawnNo3 = int(cols[4])
                marksix.DrawnNo4 = int(cols[5])
                marksix.DrawnNo5 = int(cols[6])
                marksix.DrawnNo6 = int(cols[7])
                marksix.ExtraNum = int(cols[8].split('\n')[0])
            else:
                marksix.DrawNo = cols[0].split(' ')[0][2:4]+'-'+cols[0].split(' ')[0][4:7]
                marksix.Date = datetime.strptime(cols[0].split(' ')[1], "%Y/%m/%d")

                marksix.DrawnNo1 = int(cols[1])
                marksix.DrawnNo2= int(cols[2])
                marksix.DrawnNo3 = int(cols[3])
                marksix.DrawnNo4 = int(cols[4])
                marksix.DrawnNo5 = int(cols[5])
                marksix.DrawnNo6 = int(cols[6])
                marksix.ExtraNum = int(cols[7].split('\n')[0])


            try:
                print(marksix.DrawNo, marksix.Date, marksix.DrawnNo1, marksix.DrawnNo2, marksix.DrawnNo3, marksix.DrawnNo4, marksix.DrawnNo5, marksix.DrawnNo6, marksix.ExtraNum)
                session.add(marksix)
                session.commit()
            except:
                raise
