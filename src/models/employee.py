from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class Employee(BaseModel):
    employee_name: str
    employee_id: Optional[str] = None
    action_type: Optional[str] = None
    gender: Optional[str] = None
    function: Optional[str] = None
    grade: Optional[str] = None
    job_location: Optional[str] = None
    date_of_joining: Optional[datetime] = None
    action_date: Optional[datetime] = None
