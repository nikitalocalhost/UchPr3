from sqlalchemy import Column, VARCHAR, Integer, ForeignKey

from .base import BaseModel

class Teacher(BaseModel):
    __tablename__ = 'teachers'

    name = Column(VARCHAR(255), nullable=False)
