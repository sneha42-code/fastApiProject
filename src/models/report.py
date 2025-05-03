from pydantic import BaseModel
from typing import Dict, List, Optional

class ReportSection(BaseModel):
    title: str
    content: Optional[str] = None
    charts: List[str] = []
    tables: List[Dict] = []

class ReportMetadata(BaseModel):
    file_id: str
    report_time: str
    total_employees: int
    total_exits: int
    attrition_rate: float

class Report(BaseModel):
    metadata: ReportMetadata
    sections: List[ReportSection]
