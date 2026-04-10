from fastapi import APIRouter, HTTPException, status
from datetime import datetime
import os

from models.requestmodel import GenerateDocumentRequest, GeneratePDFRequest
from models.responsemodel import GenerateDocumentResponse, ErrorResponse, SharePointMetadata, GeneratePDFResponse

from app.services.sharepoint import SharePointUtils
from app.services.documentprocessor import DocumentProcessor

router = APIRouter()


@router.post(
    "/generatedocument",
    response_model=GenerateDocumentResponse,
    summary="Generate or update Word document",
    description="Generate a new document from template or update existing document with placeholders and charts",
)
async def generate_document(request: GenerateDocumentRequest) -> GenerateDocumentResponse:
    try:
        sharepoint = SharePointUtils()
        doc_processor = DocumentProcessor()

        if request.documentIsOld == 0:
            if not request.driveId:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="driveId is required for new documents",
                )

            document_stream = sharepoint.get_document_by_name(
                request.documentName,
                is_old_document=False,
                drive_id=request.driveId,
            )
            file_name = os.path.basename(request.documentName)
            is_new_document = True

        elif request.documentIsOld == 1:
            document_stream = sharepoint.get_document_by_name(request.documentName, is_old_document=True)
            file_name = os.path.basename(request.documentName)
            is_new_document = False

        else:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="documentIsOld must be 0 (new) or 1 (existing)",
            )

        table_data = None
        if request.data:
            table_data = {
                "tag": request.data.tag,
                "headers": request.data.headers,
                "rows": request.data.rows,
                "colors": request.data.colors,
                "legend": request.data.legend,
                "headerColor": request.data.headerColor,
            }

        deployment_tables = None
        if request.deploymentTables:
            deployment_tables = []
            total = len(request.deploymentTables)
            for table_item in request.deploymentTables:
                try:
                    tbl = {
                        "tag": table_item.data.tag,
                        "headers": table_item.data.headers,
                        "rows": table_item.data.rows,
                        "colors": table_item.data.colors,
                        "legend": table_item.data.legend,
                        "headerColor": table_item.data.headerColor,
                        "isDeployment": True,
                        "deploymentTableCount": total,
                    }
                    deployment_tables.append(tbl)
                except Exception:
                    continue

        project_brief_data = None
        if request.projectBrief:
            project_brief_data = {
                "tag": request.projectBrief.tag,
                "items": request.projectBrief.items,
            }

        processed_document = doc_processor.process_document(
            document_stream,
            request.placeholders,
            None,
            table_data,
            project_brief_data,
            deployment_tables,
        )

        if is_new_document:
            upload_response = sharepoint.upload_new_file(processed_document, file_name)
        else:
            upload_response = sharepoint.update_existing_file(request.documentId, processed_document)

        metadata = sharepoint.extract_metadata(upload_response)

        return GenerateDocumentResponse(
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
                lastModified=metadata["lastModified"],
            ),
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Document processing failed: {str(e)}",
        )


@router.post(
    "/generatepdf",
    response_model=GeneratePDFResponse,
    summary="Convert DOCX to PDF and upload to SharePoint",
    description="Convert DOCX to PDF using Microsoft Graph and upload PDF back to SharePoint",
)
async def generate_pdf(request: GeneratePDFRequest) -> GeneratePDFResponse:
    try:
        sharepoint = SharePointUtils()

        if not request.documentName:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="documentName is required (e.g., '/Templates/ANP_PSL_CPMC_R1.docx')",
            )

        if not request.driveId:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="driveId is required",
            )

        if not request.fileName:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="fileName is required (output PDF name without .pdf extension)",
            )

        try:
            pdf_stream = sharepoint.convert_docx_to_pdf_with_graph(
                request.documentName,
                request.driveId,
            )
        except Exception as conversion_error:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"DOCX to PDF conversion failed: {str(conversion_error)}",
            )

        try:
            pdf_filename = f"{request.fileName}.pdf"
            upload_response = sharepoint.upload_new_file(
                pdf_stream,
                pdf_filename,
                folder_path="/Output",
            )
        except Exception as upload_error:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to upload PDF to SharePoint: {str(upload_error)}",
            )

        try:
            metadata = sharepoint.extract_metadata(upload_response)
            return GeneratePDFResponse(
                status="success",
                message="DOCX successfully converted to PDF and uploaded to SharePoint",
                documentName=metadata["fileName"],
                sharepointUrl=metadata["webUrl"],
                fileId=metadata["fileId"],
                size=metadata["size"],
                processedAt=datetime.utcnow().isoformat(),
            )
        except Exception as metadata_error:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to extract upload metadata: {str(metadata_error)}",
            )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"PDF generation failed: {str(e)}",
        )


@router.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    return ErrorResponse(
        status="failure",
        message=exc.detail,
        error_code=f"HTTP_{exc.status_code}",
    )
