from pydantic import BaseModel, Field
from typing import Optional
from datetime import datetime

class SharePointMetadata(BaseModel):
    fileId: str = Field(..., description="SharePoint file ID")
    fileName: str = Field(..., description="Name of the generated file")
    webUrl: str = Field(..., description="SharePoint web URL to access the file")
    version: int = Field(..., description="Version number of the document")
    size: int = Field(..., description="File size in bytes")
    lastModified: str = Field(..., description="Last modified timestamp")

class GenerateDocumentResponse(BaseModel):
    status: str = Field(..., description="Operation status: 'success' or 'failure'")
    message: str = Field(..., description="Human-readable message about the operation")
    documentId: Optional[str] = Field(None, description="SharePoint document ID")
    version: Optional[int] = Field(None, description="Document version number")
    sharepointUrl: Optional[str] = Field(None, description="Direct URL to SharePoint document")
    processedAt: str = Field(default_factory=lambda: datetime.utcnow().isoformat(), description="Processing completion timestamp")
    metadata: Optional[SharePointMetadata] = Field(None, description="Detailed SharePoint file metadata")
    
    class Config:
        json_schema_extra = {
            "example": {
                "status": "success",
                "message": "Document generated successfully",
                "documentId": "DOC123456",
                "version": 1,
                "sharepointUrl": "https://company.sharepoint.com/sites/docs/file123.docx",
                "processedAt": "2025-11-04T10:30:00Z",
                "metadata": {
                    "fileId": "DOC123456",
                    "fileName": "Generated_Report_20251104.docx",
                    "webUrl": "https://company.sharepoint.com/sites/docs/file123.docx",
                    "version": 1,
                    "size": 245760,
                    "lastModified": "2025-11-04T10:30:00Z"
                }
            }
        }

class ErrorResponse(BaseModel):
    status: str = Field(default="failure", description="Operation status")
    message: str = Field(..., description="Error message")
    error_code: Optional[str] = Field(None, description="Specific error code")
    processedAt: str = Field(default_factory=lambda: datetime.utcnow().isoformat(), description="Error timestamp")
    
    class Config:
        json_schema_extra = {
            "example": {
                "status": "failure",
                "message": "Template not found in SharePoint",
                "error_code": "TEMPLATE_NOT_FOUND",
                "processedAt": "2025-11-04T10:30:00Z"
            }
        }
