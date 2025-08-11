from sqlalchemy import create_engine, Column, String, Integer
from sqlalchemy.orm import sessionmaker, declarative_base
from path import *

db = create_engine(f'sqlite:///{BASE_DIR}/DB/mybd.db')
Session = sessionmaker(bind=db)
session = Session()

Base = declarative_base()

class Pessoa(Base):
    __tablename__ = 'pessoas'
    id = Column('id', Integer, autoincrement=True, primary_key=True)
    cod = Column('cod', String)
    # nome = Column('nome', String)
    cc = Column('cc', Integer)
    dia = Column('dia', String)
    horario = Column('horario', String)
    
    def __init__(self, cod, cc, dia, horario):
        self.cod = cod
        # self.nome = nome
        self.cc = cc
        self.dia = dia
        self.horario = horario

Base.metadata.create_all(bind=db)