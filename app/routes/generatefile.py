from fastapi import APIRouter, HTTPException, status
from typing import Dict, Any
from datetime import datetime
import io
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models.requestmodel import GenerateDocumentRequest
from models.responsemodel import GenerateDocumentResponse, ErrorResponse, SharePointMetadata

from services.sharepoint import SharePointUtils
from services.documentprocessor import DocumentProcessor
from services.imageservice import generate_chart_image

router = APIRouter()

@router.post("/generatedocument", 
             response_model=GenerateDocumentResponse,
             summary="Generate or update Word document",
             description="Generate a new document from template or update existing document with placeholders and charts")
async def generate_document(request: GenerateDocumentRequest) -> GenerateDocumentResponse:
    
    try:
        sharepoint = SharePointUtils()
        doc_processor = DocumentProcessor()

        if request.documentIsOld == 0:
            if not request.driveId:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="driveId is required for new documents"
                )

            document_stream = sharepoint.get_document_by_name(request.documentName, is_old_document=False, drive_id=request.driveId)
            file_name = os.path.basename(request.documentName)
            is_new_document = True

        elif request.documentIsOld == 1:
            document_stream = sharepoint.get_document_by_name(request.documentName, is_old_document=True)
            file_name = os.path.basename(request.documentName)
            is_new_document = False

        else:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="documentIsOld must be 0 (new) or 1 (existing)"
            )

        chart_images: Dict[str, io.BytesIO] = {}
        
        if request.charts:
            for chart in request.charts:
                try:
                    if chart.chartType.lower() in ["table", "dynamic_table"]:
                        chart_image = generate_chart_image(chart.data, chart.title)
                        chart_images[chart.tag] = chart_image
                    else:
                        chart_image = generate_chart_image(chart.data, chart.title)
                        chart_images[chart.tag] = chart_image
                        
                except Exception as e:
                    continue

        table_data = None
        if request.data:
            table_data = {
                "tag": request.data.tag,
                "headers": request.data.headers,
                "rows": request.data.rows,
                "colors": request.data.colors,
                "legend": request.data.legend,
                "headerColor": request.data.headerColor
            }
        
        processed_document = doc_processor.process_document(
            document_stream, 
            request.placeholders, 
            chart_images,
            table_data
        )
        if is_new_document:
            upload_response = sharepoint.upload_new_file(processed_document, file_name)
        else:
            upload_response = sharepoint.update_existing_file(request.documentId, processed_document)
        
        metadata = sharepoint.extract_metadata(upload_response)
        
        response = GenerateDocumentResponse(
            status="success",
            message="Document generated successfully" if is_new_document else "Document updated successfully",
            documentId=metadata["fileId"],
            version=metadata["version"],
            sharepointUrl=metadata["webUrl"],
            processedAt=datetime.utcnow().isoformat(),
            metadata=SharePointMetadata(
                fileId=metadata["fileId"],
                fileName=metadata["fileName"],
                webUrl=metadata["webUrl"],
                version=metadata["version"],
                size=metadata["size"],
                lastModified=metadata["lastModified"]
            )
        )
        
        return response
        
    except HTTPException as http_exc:
        raise

    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Document processing failed: {str(e)}"
        )

@router.get("/health", summary="Health check endpoint")
async def health_check():
    """Simple health check endpoint"""
    return {"status": "healthy", "timestamp": datetime.utcnow().isoformat()}

@router.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    return ErrorResponse(
        status="failure",
        message=exc.detail,
        error_code=f"HTTP_{exc.status_code}"
    )
