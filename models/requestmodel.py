from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional, Union

class DynamicTableData(BaseModel):
    tag: str = Field(default="table", description="Tag name of content control where table should be inserted (default: 'table')")
    headers: Optional[List[str]] = Field(None, description="Table headers (optional)")
    rows: List[Dict[str, Any]] = Field(..., description="Array of row data objects for table generation")
    colors: Optional[List[Dict[str, str]]] = Field(None, description="Array of color mappings for phases/categories")
    legend: Optional[List[Dict[str, str]]] = Field(None, description="Array of legend items with phase and color")
    headerColor: Optional[str] = Field("#333399", description="Header row color (default: #333399)")
    
class RequestMetadata(BaseModel):
    generatedBy: Optional[str] = Field(None, description="User who generated the document")
    requestedAt: Optional[str] = Field(None, description="Timestamp when request was made")
    purpose: Optional[str] = Field(None, description="Purpose of document generation")

class GenerateDocumentRequest(BaseModel):
    documentIsOld: int = Field(..., description="0 for new document, 1 for existing document update")
    documentName: str = Field(..., description="Document filename (required for both new and existing documents)")
    placeholders: Dict[str, Union[str, int, float]] = Field(..., description="Key-value pairs for text replacement")
    data: Optional[DynamicTableData] = Field(None, description="Dynamic table data to generate and insert into document")
    meta: Optional[RequestMetadata] = Field(None, description="Request metadata")
    
    class Config:
        json_schema_extra = {
            "example": {
                "documentIsOld": 0,
                "documentName": "ANP_PSL_CPMC_R1_Template.docx",
                "placeholders": {
                    "CompanyName": "Vertex Constructions LLP",
                    "CreatedDate": "08 January 2026",
                    "ProposalCode": "VC-CPMC-026",
                    "CompanyAddress": "Mr. Arjun Mehta, /n45 Tech Park Road, /nHyderabad, Telangana",
                    "ProjectName": "Skyline Business Hub",
                    "Subject": "CONSTRUCTION PROJECT MANAGEMENT CONSULTANCY SERVICES",
                    "Reference": "Discussion held on 05 January 2026",
                    "Regards": "SURESH K R,",
                    "TotalBUAinSqm": "3,750",
                    "DurationinMonths": "18",
                    "BuiltEnvironment": "Mixed-use Commercial Development",
                    "NumberofFloors": "2B + G + 5 Floors",
                    "CPMCServices": "Project Planning, Cost Control, Contract Administration, Quality Assurance & Safety Management",
                    "ProfessionalFee": "Rs. 22.50 Lakhs (Rupees Twenty Two Lakhs and Fifty Thousand only)",
                    "ProjectDuration": "12",
                    "ProjectLocation": "OfficE Park",
                    "PlanningPhase": "2",
                    "ConstructionPhase": "12",
                    "CloseoutPhase": "4",
                    "StaffDeployment": "Quantity Surveyor"
                },
                "data": {
                    "tag": "Table",
                    "headers": ["Sl. No.", "Staff", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Total"],
                    "rows": [
                        {
                            "Sl. No.": "1",
                            "Staff": "Project Engineer",
                            "months": [
                                {"Dec-25": {"phase": "Planning Phase", "value": "1"}},
                                {"Jan-26": {"phase": "Planning Phase", "value": "1"}},
                                {"Feb-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Mar-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Apr-26": {"phase": "Construction Phase", "value": "1"}},
                                {"May-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Jun-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Jul-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Aug-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Sep-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Oct-26": {"phase": "Closeout Phase", "value": "1"}},
                                {"Nov-26": {"phase": "Closeout Phase", "value": "1"}}
                            ],
                            "Total": "12"
                        },
                        {
                            "Sl. No.": "2",
                            "Staff": "Site Supervisor",
                            "months": [
                                {"Dec-25": {"phase": "Planning Phase", "value": "0"}},
                                {"Jan-26": {"phase": "Planning Phase", "value": "0"}},
                                {"Feb-26": {"phase": "Construction Phase", "value": "2"}},
                                {"Mar-26": {"phase": "Construction Phase", "value": "2"}},
                                {"Apr-26": {"phase": "Construction Phase", "value": "2"}},
                                {"May-26": {"phase": "Construction Phase", "value": "2"}},
                                {"Jun-26": {"phase": "Construction Phase", "value": "2"}},
                                {"Jul-26": {"phase": "Construction Phase", "value": "2"}},
                                {"Aug-26": {"phase": "Construction Phase", "value": "2"}},
                                {"Sep-26": {"phase": "Construction Phase", "value": "2"}},
                                {"Oct-26": {"phase": "Closeout Phase", "value": "2"}},
                                {"Nov-26": {"phase": "Closeout Phase", "value": "2"}}
                            ],
                            "Total": "20"
                        },
                        {
                            "Sl. No.": "3",
                            "Staff": "Quantity Surveyor",
                            "months": [
                                {"Dec-25": {"phase": "Planning Phase", "value": "1"}},
                                {"Jan-26": {"phase": "Planning Phase", "value": "1"}},
                                {"Feb-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Mar-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Apr-26": {"phase": "Construction Phase", "value": "1"}},
                                {"May-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Jun-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Jul-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Aug-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Sep-26": {"phase": "Construction Phase", "value": "1"}},
                                {"Oct-26": {"phase": "Closeout Phase", "value": "0"}},
                                {"Nov-26": {"phase": "Closeout Phase", "value": "0"}}
                            ],
                            "Total": "10"
                        }
                    ],
                    "colors": [
                        {"Planning Phase": "#FFC000"},
                        {"Construction Phase": "#00B050"},
                        {"Closeout Phase": "#4472C4"}
                    ],
                    "legend": [
                        {"phase": "Planning Phase", "color": "#FFC000"},
                        {"phase": "Construction Phase", "color": "#00B050"},
                        {"phase": "Closeout Phase", "color": "#4472C4"}
                    ],
                    "headerColor": "#333399"
                }
            }
        }

