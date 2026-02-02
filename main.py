from fastapi import FastAPI, HTTPException, status
from fastapi import Depends
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
import jwt
from fastapi.middleware.cors import CORSMiddleware
from datetime import datetime
from models.requestmodel import GenerateDocumentRequest
from models.responsemodel import GenerateDocumentResponse, SharePointMetadata
import uvicorn

app = FastAPI(
    title="Document Generator API",
    description="API for generating and updating Word documents with dynamic content from SharePoint templates",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],  
    allow_headers=["*"], 
)

JWT_ISSUER = "https://localhost:7153/"
JWT_AUDIENCE = "https://localhost:7153/"
JWT_SECRET = "M2vAjdN7XqK8cFpZ9sYTwuRZB3HLqVnJxG0btDm4EyUV1WCkhrfTa5g6MzQeLSnP&&"
JWT_ALGORITHM = "HS256"
JWT_EXP_MINUTES = 10

security = HTTPBearer()

def verify_jwt(credentials: HTTPAuthorizationCredentials = Depends(security)):
    token = credentials.credentials
    try:
        payload = jwt.decode(
            token,
            JWT_SECRET,
            algorithms=[JWT_ALGORITHM],
            audience=JWT_AUDIENCE,
            issuer=JWT_ISSUER,
            options={"require": ["exp", "iss", "aud"]}
        )
        return payload
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except jwt.InvalidTokenError as e:
        raise HTTPException(status_code=401, detail=f"Invalid token: {str(e)}")

@app.post("/api/generatedocument", 
         response_model=GenerateDocumentResponse,
         summary="Generate or update Word document",
         description="Generate a new document from template or update existing document with placeholders and dynamic tables")
async def generate_document(request: GenerateDocumentRequest, token_payload: dict = Depends(verify_jwt)) -> GenerateDocumentResponse:
    try:
        from app.services.sharepoint import SharePointUtils
        from app.services.documentprocessor import DocumentProcessor
        
        sharepoint = SharePointUtils()
        doc_processor = DocumentProcessor()
        
        if not request.documentName:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="documentName is required for both new and existing documents"
            )
        
        is_existing_document = request.documentIsOld == 1 or request.documentIsOld is True
        
        try:
            if is_existing_document:
                document_stream = sharepoint.get_document_by_name(request.documentName, is_old_document=True)
                file_name = request.documentName
                current_version = 1
            else:
                if not request.driveId:
                    raise HTTPException(
                        status_code=status.HTTP_400_BAD_REQUEST,
                        detail="driveId is required for new documents"
                    )
                document_stream = sharepoint.get_document_by_name(request.documentName, is_old_document=False, drive_id=request.driveId)
                file_name = sharepoint.generate_file_name("Report")
                current_version = 0
        except Exception as e:
            location = "Output folder" if is_existing_document else "specified drive"
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Document '{request.documentName}' not found in {location}: {str(e)}"
            )
        
        table_data = None
        if request.data:
            table_data = {
                "tag": getattr(request.data, "tag", "table"),
                "headers": getattr(request.data, "headers", None),
                "rows": request.data.rows,
                "colors": getattr(request.data, "colors", None),
                "legend": getattr(request.data, "legend", None),
                "headerColor": getattr(request.data, "headerColor", "#333399")
            }
        
        project_brief_data = None
        if request.projectBrief:
            project_brief_data = {
                "tag": getattr(request.projectBrief, "tag", "ProjectBrief"),
                "items": request.projectBrief.items
            }
        
        try:
            processed_document = doc_processor.process_document(
                document_stream,
                request.placeholders,
                None, 
                table_data,
                project_brief_data
            )
        except Exception as processing_error:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Document processing failed: {str(processing_error)}"
            )
        
        try:
            if is_existing_document:
                upload_response = sharepoint.upload_new_file(
                    processed_document, 
                    request.documentName, 
                    folder_path="/Output"
                )
                new_version = current_version + 1
            else:
                upload_response = sharepoint.upload_new_file(processed_document, file_name, folder_path="/Output")
                new_version = 1
        except Exception as upload_error:
            error_message = str(upload_error)
            
            if "resourceLocked" in error_message or "locked" in error_message.lower():
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    base_name = file_name.replace(".docx", "")
                    fallback_name = f"{base_name}_updated_{timestamp}.docx"
                    upload_response = sharepoint.upload_new_file(processed_document, fallback_name, folder_path="/Output")
                    new_version = 1
                except Exception as fallback_error:
                    raise HTTPException(
                        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                        detail=f"Document is locked and fallback creation failed: {str(fallback_error)}"
                    )
            else:
                raise HTTPException(
                    status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                    detail=f"Failed to save document to SharePoint: {str(upload_error)}"
                )
        
        metadata = sharepoint.extract_metadata(upload_response)
        metadata["version"] = new_version
        
        placeholders_count = len(request.placeholders) if request.placeholders else 0
        has_table = request.data is not None
        
        if is_existing_document:
            success_message = f"Document updated successfully (version {current_version} → {new_version})"
        else:
            success_message = f"New document created successfully from template"
        
        details = []
        if placeholders_count > 0:
            details.append(f"{placeholders_count} placeholders replaced")
        if has_table:
            details.append(f"dynamic table inserted")
        
        if details:
            success_message += f" - {', '.join(details)}"
        
        return GenerateDocumentResponse(
            status="success",
            message=success_message,
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
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Document processing failed: {str(e)}"
        )

@app.get("/")
async def root():
    return {
        "message": "Document Generator API", 
        "status": "running",
        "docs": "/docs",
        "generate_endpoint": "/api/generatedocument"
    }

@app.get("/health")
def health():
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
