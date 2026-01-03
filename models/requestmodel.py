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
    driveId: Optional[str] = Field(None, description="Drive ID for new documents (required when documentIsOld=0)")
    placeholders: Dict[str, Union[str, int, float]] = Field(..., description="Key-value pairs for text replacement")
    data: Optional[DynamicTableData] = Field(None, description="Dynamic table data to generate and insert into document")
    meta: Optional[RequestMetadata] = Field(None, description="Request metadata")
    
    class Config:
        json_schema_extra = {
            "example": {
                "data": {
                    "tag": "Table",
                    "headerColor": "#333399",
                    "headers": [
                        "Sl. No.",
                        "Staff",
                        "Dec-25",
                        "Jan-26",
                        "Feb-26",
                        "Mar-26",
                        "Apr-26",
                        "May-26",
                        "Jun-26",
                        "Jul-26",
                        "Aug-26",
                        "Sep-26",
                        "Oct-26",
                        "Nov-26",
                        "Total"
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
                    "rows": [
                        {
                            "Sl. No.": "1",
                            "Staff": "Project Engineer",
                            "Total": "12",
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
                            ]
                        },
                        {
                            "Sl. No.": "2",
                            "Staff": "Site Supervisor",
                            "Total": "20",
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
                            ]
                        },
                        {
                            "Sl. No.": "3",
                            "Staff": "Quantity Surveyor",
                            "Total": "10",
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
                            ]
                        }
                    ]
                },
                "documentIsOld": 0,
                "driveId": "b!jtW2losKJ0CuA4Ta-98ieMStCYTPWNlFitOnB1A_LQ7uK4_iny_SQI5e_PT_VePY",
                "documentName": "/BD-DENEC-562-2025-2026-R0.docx",
                "placeholders": {
                    "CompanyName": "Vertex Constructions LLP",
                    "CompanyAddress": "Mr. Arjun Mehta,\n45 Tech Park Road,\nHyderabad, Telangana",
                    "ProjectName": "Skyline Business Hub",
                    "ProjectLocation": "Office Park",
                    "BuiltEnvironment": "Mixed-use Commercial Development",
                    "NumberofFloors": "2B + G + 5 Floors",
                    "TotalBUAinSqm": "3,750",
                    "DurationinMonths": "18",
                    "PlanningPhase": "2",
                    "ConstructionPhase": "12",
                    "CloseoutPhase": "4",
                    "ProjectDuration": "12",
                    "CPMCServices": "Project Planning, Cost Control, Contract Administration, Quality Assurance & Safety Management",
                    "ProfessionalFee": "Rs. 22.50 Lakhs (Rupees Twenty Two Lakhs and Fifty Thousand only)",
                    "ProposalCode": "VC-CPMC-026",
                    "Reference": "Discussion held on 05 January 2026",
                    "CreatedDate": "08 January 2026",
                    "Subject": "CONSTRUCTION PROJECT MANAGEMENT CONSULTANCY SERVICES",
                    "StaffDeployment": "Quantity Surveyor",
                    "Regards": "SURESH K R,"
                }
            }
        }

