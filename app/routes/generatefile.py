from fastapi import APIRouter, HTTPException, status
from typing import Dict, Any
import logging
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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create router
router = APIRouter()

@router.post("/generatedocument", 
             response_model=GenerateDocumentResponse,
             summary="Generate or update Word document",
             description="Generate a new document from template or update existing document with placeholders and charts")
async def generate_document(request: GenerateDocumentRequest) -> GenerateDocumentResponse:
    
    try:
        sharepoint = SharePointUtils()
        doc_processor = DocumentProcessor()
        
        logger.info(f"Processing document request - documentIsOld: {request.documentIsOld}")
        
        if request.documentIsOld == 0:
            if not request.templateId:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="templateId is required for new documents"
                )
            
            logger.info(f"Fetching template: {request.templateId}")
            document_stream = sharepoint.get_template_by_id(request.templateId)
            file_name = sharepoint.generate_file_name("Report")
            is_new_document = True
            
        elif request.documentIsOld == 1:
            if not request.documentId:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="documentId is required for existing documents"
                )
            
            logger.info(f"Fetching existing document: {request.documentId}")
            document_stream = sharepoint.download_file_by_id(request.documentId)
            is_new_document = False
            
        else:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="documentIsOld must be 0 (new) or 1 (existing)"
            )
        
        chart_images: Dict[str, io.BytesIO] = {}
        
        if request.charts:
            logger.info(f"Generating {len(request.charts)} charts")
            for chart in request.charts:
                try:
                    # Generate chart image based on chartType
                    if chart.chartType.lower() in ["table", "dynamic_table"]:
                        chart_image = generate_chart_image(chart.data, chart.title)
                        chart_images[chart.tag] = chart_image
                        logger.info(f"Generated chart for tag: {chart.tag}")
                    else:
                        chart_image = generate_chart_image(chart.data, chart.title)
                        chart_images[chart.tag] = chart_image
                        logger.info(f"Generated default table chart for tag: {chart.tag}")
                        
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
        
        logger.info(f"Document processing completed successfully: {metadata['fileId']}")
        return response
        
    except HTTPException:
        # Re-raise HTTP exceptions
        raise
        
    except Exception as e:
        logger.error(f"Document processing failed: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Document processing failed: {str(e)}"
        )

@router.get("/health", summary="Health check endpoint")
async def health_check():
    """Simple health check endpoint"""
    return {"status": "healthy", "timestamp": datetime.utcnow().isoformat()}

# Error handlers
@router.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    return ErrorResponse(
        status="failure",
        message=exc.detail,
        error_code=f"HTTP_{exc.status_code}"
    )
