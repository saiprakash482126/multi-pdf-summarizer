from fastapi import FastAPI, HTTPException, UploadFile, File, BackgroundTasks, Header, Form, Request, Depends
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, HttpUrl, Field
from typing import Optional, List, Dict, Any, Union
import os
import tempfile
import shutil
from pathlib import Path
from datetime import datetime
import uuid
import json
import requests
from urllib.parse import quote
import jwt
from jwt import PyJWKClient, PyJWTError, ExpiredSignatureError
from dotenv import load_dotenv
from dataclasses import dataclass
from odd_check_analyzer import analyze_document as analyze_pharma_file
from odd_check_analyzer import analyze_document as analyze_pharmaceutical_content


# Load environment variables
load_dotenv()

# Import your modules
from document_validator import DocumentValidator, process_document
from summarize_onedrive import process_onedrive_folder

app = FastAPI(
    title="Document Processing API",
    description="API for document validation and summarization",
    version="1.0.0"
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Request/Response models
class ValidationRequest(BaseModel):
    document_url: Optional[str] = None
    document_text: Optional[str] = None

class ValidationResponse(BaseModel):
    is_valid: bool
    issues: List[Dict[str, Any]] = []
    summary: Optional[str] = None
    file_name: Optional[str] = None

class SummarizeRequest(BaseModel):
    folder_path: str
    recursive: bool = False

class SummarizeResponse(BaseModel):
    job_id: str
    status: str
    message: str
    summary: Optional[Dict[str, Any]] = None

class OneDriveValidationRequest(BaseModel):
    """Request model for OneDrive folder validation."""
    root_folder_name: str
    recursive: bool = True
    user_id: Optional[str] = None  # Made optional as we'll get it from the JWT token

class FileValidationResult(BaseModel):
    """Model for individual file validation result."""
    file_path: str
    is_valid: bool
    issues: List[Dict[str, Any]] = []

class OneDriveValidationResponse(BaseModel):
    """Response model for OneDrive folder validation."""
    status: str
    processed_files: int
    valid_files: int
    invalid_files: int
    results: List[FileValidationResult] = []
    message: Optional[str] = None




class PharmaceuticalAnalysisResponse(BaseModel):
    status: str
    file: Optional[str] = None
    entities: Dict[str, List[str]] = {}
    text_sample: Optional[str] = None
    error: Optional[str] = None

# Add ValidationErrorDetail dataclass
@dataclass
class ValidationErrorDetail:
    """Class for storing validation error details."""
    folder_path: str
    file_name: str
    error_code: str
    status: str
    message: str
    details: str
    error_type: str


# JWT Configuration
JWT_SECRET = os.getenv("JWT_SECRET", "5367566B59703373367639792F423F4528482B4D6251655468576D5A71347437")
JWT_ALGORITHM = "HS256"

def get_onedrive_user_id(request: Request) -> str:
    """
    Extract the OneDrive user ID from the JWT token in the Authorization header.
    
    Args:
        request: The FastAPI request object
        
    Returns:
        str: The OneDrive user ID from the token
        
    Raises:
        HTTPException: If the token is invalid or missing the OneDrive user ID
    """
    try:
        print("\n=== Starting token validation ===")
        
        # Get the Authorization header
        auth_header = request.headers.get("Authorization")
        print(f"Auth header: {'Present' if auth_header else 'Missing'}")
        
        if not auth_header or not auth_header.startswith("Bearer "):
            raise HTTPException(
                status_code=401,
                detail="Missing or invalid Authorization header. Expected format: 'Bearer <token>'"
            )
        
        # Extract the token
        token = auth_header.split(" ")[1].strip()
        print("Token extracted successfully")
        
        try:
            # First, decode without verification to check the algorithm
            unverified_header = jwt.get_unverified_header(token)
            print(f"Token header: {unverified_header}")
            
            algorithm = unverified_header.get('alg')
            if not algorithm:
                raise HTTPException(
                    status_code=401,
                    detail="Token is missing algorithm in header"
                )
                
            print(f"Token algorithm: {algorithm}")
            
            # List of supported algorithms (expanded list)
            supported_algorithms = ["RS256", "HS256", "RS384", "RS512", "ES256", "ES384", "ES512"]
            
            # For debugging - log the algorithm even if not supported
            if algorithm not in supported_algorithms:
                print(f"Warning: Token uses unsupported algorithm: {algorithm}")
                print("Attempting to decode anyway...")
            
            # Get the key based on the algorithm
            key = JWT_SECRET  # Default to HS256 secret
            
            # Handle RS* and ES* algorithms (asymmetric)
            if algorithm.startswith(('RS', 'ES')):
                jwks_url = os.getenv("JWKS_URL")
                if not jwks_url:
                    print("Warning: JWKS_URL not set in environment variables")
                    # If we can't verify, try to decode without verification for debugging
                    print("Attempting to decode without verification for debugging...")
                    payload = jwt.decode(
                        token,
                        options={"verify_signature": False}
                    )
                    print(f"Debug - Unverified token payload: {json.dumps(payload, indent=2)}")
                    raise HTTPException(
                        status_code=401,
                        detail=(
                            f"Cannot verify token with algorithm {algorithm}. "
                            "JWKS_URL not configured. Please set JWKS_URL in your environment variables."
                        )
                    )
                
                print(f"Using JWKS URL: {jwks_url}")
                try:
                    jwks_client = PyJWKClient(jwks_url)
                    signing_key = jwks_client.get_signing_key_from_jwt(token)
                    key = signing_key.key
                except Exception as jwks_error:
                    print(f"Error fetching JWKS: {str(jwks_error)}")
                    raise HTTPException(
                        status_code=401,
                        detail=f"Failed to fetch verification key: {str(jwks_error)}"
                    )
            
            # Decode the JWT token with the appropriate key
            print(f"Decoding token with algorithm: {algorithm}")
            try:
                payload = jwt.decode(
                    token,
                    key=key,
                    algorithms=supported_algorithms,  # Pass all supported algorithms
                    options={
                        "verify_signature": True,
                        "verify_exp": True,
                        "verify_aud": False,  # You might want to enable this in production
                        "verify_iss": False   # You might want to enable this in production
                    }
                )
            except jwt.InvalidAlgorithmError as iae:
                # If we get here, the algorithm is not supported by PyJWT
                raise HTTPException(
                    status_code=401,
                    detail=f"Unsupported algorithm: {algorithm}. This server does not support this JWT algorithm."
                )
            
            print(f"Token payload: {json.dumps(payload, indent=2)}")
            
            # Try to get user ID from various claims
            user_claims = [
                "oid",          # Microsoft Object ID
                "sub",          # Subject
                "upn",          # User Principal Name
                "email",
                "preferred_username",
                "unique_name",
                "name",
                "email_verified",
                "onprem_sid",
                "tid"           # Tenant ID
            ]
            
            # Get all available claims for debugging
            available_claims = {claim: payload.get(claim) for claim in user_claims if claim in payload}
            print(f"Available user claims: {json.dumps(available_claims, indent=2)}")
            
            # Try to get user ID from standard claims
            onedrive_user_id = (
                payload.get("oid") or
                payload.get("sub") or
                payload.get("upn") or
                payload.get("email")
            )
            
            # Fallback to other claims if needed
            if not onedrive_user_id:
                onedrive_user_id = (
                    payload.get("preferred_username") or
                    payload.get("unique_name") or
                    payload.get("name")
                )
            
            # If still no user ID found, check for custom claims
            if not onedrive_user_id:
                custom_claims = ["onprem_sid", "tid", "azp", "appid", "client_id"]
                for claim in custom_claims:
                    if claim in payload:
                        onedrive_user_id = payload[claim]
                        print(f"Using custom claim '{claim}' as user ID")
                        break
            
            if not onedrive_user_id:
                # For debugging, include all available claims in the error
                all_claims = list(payload.keys())
                print(f"All available claims in token: {all_claims}")
                
                # Try to find any claim that might be a user identifier
                potential_ids = [
                    payload[claim] for claim in all_claims 
                    if isinstance(payload[claim], (str, int)) and 
                    str(payload[claim]).strip() and
                    len(str(payload[claim])) < 100  # Filter out very long values
                ]
                
                if potential_ids:
                    onedrive_user_id = str(potential_ids[0])
                    print(f"Using first available claim value as user ID: {onedrive_user_id}")
                else:
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            "Could not determine user identity from token. "
                            f"Available claims: {', '.join(all_claims)}"
                        )
                    )
            
            print(f"Extracted user ID: {onedrive_user_id}")
            return str(onedrive_user_id)
            
        except jwt.PyJWTError as e:
            error_msg = str(e) or "Unknown JWT error"
            print(f"JWT Error: {error_msg}")
            print(f"Token (first 50 chars): {token[:50]}...")
            
            # Try to get more details about the token
            try:
                # Try to decode without verification to see the header
                header = jwt.get_unverified_header(token)
                print(f"Token header: {header}")
                
                # Try to decode payload without verification for debugging
                try:
                    payload = jwt.decode(token, options={"verify_signature": False})
                    print(f"Token payload (unverified): {json.dumps(payload, indent=2)}")
                except Exception as payload_error:
                    print(f"Could not decode token payload: {str(payload_error)}")
                    
            except Exception as header_error:
                print(f"Could not decode token header: {str(header_error)}")
            
            raise HTTPException(
                status_code=401,
                detail=f"Invalid token: {error_msg}",
                headers={"WWW-Authenticate": f"Bearer error=\"invalid_token\", error_description=\"{error_msg}\""}
            )
            
    except HTTPException:
        raise
        
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"\n=== UNEXPECTED ERROR ===\n{error_details}\n======================")
        
        # Include more context in the error response
        context = {
            "error": str(e) or "Unknown error",
            "type": type(e).__name__,
            "traceback": error_details.split('\n')  # Send as array for better readability
        }
        
        raise HTTPException(
            status_code=500,
            detail={
                "message": "An unexpected error occurred while processing the token",
                "details": context
            }
        )

# Store background tasks
task_status = {}

async def get_request_body(request: Request):
    try:
        return await request.json()
    except:
        return None

def get_user_identifier_from_token(token: str) -> str:
    """Extract user identifier from JWT token (works with both app and delegated tokens)"""
    try:
        # Decode without verification for testing
        decoded = jwt.decode(
            token,
            options={"verify_signature": False, "verify_aud": False, "verify_exp": False}
        )
        
        print("Token claims:", json.dumps(decoded, indent=2))  # Debug: print all claims
        
        # For application tokens, we can use the oid (object ID) to identify the user
        if decoded.get('idtyp') == 'app':
            print("Using application token with OID")
            user_oid = decoded.get('oid')
            if not user_oid:
                raise ValueError("Application token is missing 'oid' claim")
            return user_oid
        
        # For delegated tokens, try to get UPN/email
        user_identifier = decoded.get('upn') or decoded.get('email') or decoded.get('preferred_username')
        if not user_identifier:
            raise ValueError("Could not find user identifier in token")
            
        print(f"Found user identifier in token: {user_identifier}")
        return user_identifier
        
    except Exception as e:
        print(f"Error decoding token: {str(e)}")
        if 'decoded' in locals():
            print("Token contents:", decoded)
        raise HTTPException(status_code=400, detail=f"Invalid access token: {str(e)}")

# @app.post("/validate", response_model=ValidationResponse)
# async def validate_document(
#     request: Request,
#     file: Optional[UploadFile] = File(None),
#     document_url: Optional[str] = Form(None),
#     document_text: Optional[str] = Form(None)
# ):
#     """
#     Validate a document from file upload, URL, or direct text.
#     Supports both JSON and form-data requests.
    
#     JSON Request Example:
#     {
#         "document_url": "https://example.com/document.pdf"
#         // or
#         "document_text": "Document content here..."
#     }
    
#     Form-Data Request Example:
#     - file: [file content]
#     - document_url: "https://example.com/document.pdf"
#     - document_text: "Document content here..."
#     """
#     try:
#         # Check if request is JSON
#         content_type = request.headers.get('content-type', '')
        
#         if 'application/json' in content_type:
#             # Handle JSON request
#             json_data = await request.json()
#             document_url = json_data.get('document_url')
#             document_text = json_data.get('document_text')
#             file = None
        
#         # Initialize response
#         result = {
#             "is_valid": True,
#             "issues": [],
#             "file_name": None
#         }
        
#         temp_dir = tempfile.mkdtemp()
        
#         try:
#             # Handle file upload
#             if file:
#                 # Create a secure filename
#                 file_extension = os.path.splitext(file.filename or 'document')[1] or '.bin'
#                 safe_filename = f"{uuid.uuid4()}{file_extension}"
#                 file_path = os.path.join(temp_dir, safe_filename)
                
#                 # Save the uploaded file
#                 with open(file_path, "wb") as buffer:
#                     shutil.copyfileobj(file.file, buffer)
                
#                 # Process the document
#                 validation_results = process_document(file_path)
                
#                 # Convert validation results to response format
#                 issues = []
#                 for severity, items in validation_results.items():
#                     for item in items:
#                         issues.append({
#                             "severity": severity,
#                             "issue": item.get("issue", ""),
#                             "location": item.get("location", ""),
#                             "details": item.get("details", "")
#                         })
                
#                 # Update result
#                 result.update({
#                     "is_valid": not any(iss["severity"] == "Error" for iss in issues),
#                     "issues": issues,
#                     "file_name": file.filename
#                 })
                
#             # Handle URL or text input if no file was uploaded
#             elif document_url or document_text:
#                 # Create a temporary file for text content
#                 temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4()}.txt")
#                 with open(temp_file_path, "w", encoding="utf-8") as f:
#                     if document_url:
#                         f.write(f"[URL Content: {document_url}]\n\n")
#                         # In a real implementation, you would download the URL content here
#                     if document_text:
#                         f.write(document_text)
                
#                 # Process the document
#                 validation_results = process_document(temp_file_path)
                
#                 # Convert validation results to response format
#                 issues = []
#                 for severity, items in validation_results.items():
#                     for item in items:
#                         issues.append({
#                             "severity": severity,
#                             "issue": item.get("issue", ""),
#                             "location": item.get("location", ""),
#                             "details": item.get("details", "")
#                         })
                
#                 # Update result
#                 result.update({
#                     "is_valid": not any(iss["severity"] == "Error" for iss in issues),
#                     "issues": issues,
#                     "file_name": document_url or "text_content.txt"
#                 })
                
#             else:
#                 raise HTTPException(
#                     status_code=400,
#                     detail="Either 'file', 'document_url', or 'document_text' must be provided"
#                 )
                
#             return result
            
#         finally:
#             # Clean up temporary directory
#             shutil.rmtree(temp_dir, ignore_errors=True)
            
#     except HTTPException:
#         raise
#     except Exception as e:
#         raise HTTPException(
#             status_code=500,
#             detail=f"Error processing document: {str(e)}"
#         )


@app.post("/validate", response_model=ValidationResponse)
async def validate_document(
    request: Request,
    file: Optional[UploadFile] = File(None),
    document_url: Optional[str] = Form(None),
    document_text: Optional[str] = Form(None),
    folder_name: Optional[str] = Form(None)  # ← ADDED
):
    """
    Validate a document from file upload, URL, direct text, or folder name.
    """

    try:
        # Check if request is JSON
        content_type = request.headers.get('content-type', '')
        
        if 'application/json' in content_type:
            json_data = await request.json()
            document_url = json_data.get('document_url')
            document_text = json_data.get('document_text')
            folder_name = json_data.get('folder_name')  # ← ADDED
            file = None
        
        # Initialize response
        result = {
            "is_valid": True,
            "issues": [],
            "file_name": None
        }
        
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Handle file upload
            if file:
                file_extension = os.path.splitext(file.filename or 'document')[1] or '.bin'
                safe_filename = f"{uuid.uuid4()}{file_extension}"
                file_path = os.path.join(temp_dir, safe_filename)
                
                with open(file_path, "wb") as buffer:
                    shutil.copyfileobj(file.file, buffer)
                
                validation_results = process_document(file_path)
                
                issues = []
                for severity, items in validation_results.items():
                    for item in items:
                        issues.append({
                            "severity": severity,
                            "issue": item.get("issue", ""),
                            "location": item.get("location", ""),
                            "details": item.get("details", "")
                        })
                
                result.update({
                    "is_valid": not any(iss["severity"] == "Error" for iss in issues),
                    "issues": issues,
                    "file_name": file.filename
                })
                print('working')

            # 👉 Folder Handling (FIXED)
            elif folder_name:
                # Folder must exist in your backend directory
                folder_path = os.path.join(os.getcwd(), folder_name)

                if not os.path.exists(folder_path):
                    raise HTTPException(
                        status_code=400,
                        detail=f"Folder '{folder_name}' not found on server"
                    )

                all_issues = []

                for root, dirs, files in os.walk(folder_path):
                    for f in files:
                        file_path = os.path.join(root, f)
                        validation_results = process_document(file_path)

                        for severity, items in validation_results.items():
                            for item in items:
                                all_issues.append({
                                    "file": f,
                                    "severity": severity,
                                    "issue": item.get("issue", ""),
                                    "location": item.get("location", ""),
                                    "details": item.get("details", "")
                                })
                
                result.update({
                    "is_valid": not any(iss["severity"] == "Error" for iss in all_issues),
                    "issues": all_issues,
                    "file_name": folder_name
                })

            # Handle URL or text input
            elif document_url or document_text:
                temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4()}.txt")
                with open(temp_file_path, "w", encoding="utf-8") as f:
                    if document_url:
                        f.write(f"[URL Content: {document_url}]\n\n")
                    if document_text:
                        f.write(document_text)

                validation_results = process_document(temp_file_path)
                
                issues = []
                for severity, items in validation_results.items():
                    for item in items:
                        issues.append({
                            "severity": severity,
                            "issue": item.get("issue", ""),
                            "location": item.get("location", ""),
                            "details": item.get("details", "")
                        })
                
                result.update({
                    "is_valid": not any(iss["severity"] == "Error" for iss in issues),
                    "issues": issues,
                    "file_name": document_url or "text_content.txt"
                })
                
            else:
                raise HTTPException(
                    status_code=400,
                    detail="Either 'file', 'document_url', 'document_text', or 'folder_name' must be provided"
                )
                
            return result
            
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing document: {str(e)}"
        )

@app.post("/api/summarize", response_model=SummarizeResponse)
async def start_summarization(
    request: SummarizeRequest,
    background_tasks: BackgroundTasks
):
    """
    Start a background task to summarize documents in a OneDrive folder
    """
    job_id = str(uuid.uuid4())
    
    # Initialize task status
    task_status[job_id] = {
        "status": "processing",
        "message": "Task started",
        "start_time": datetime.utcnow(),
        "result": None
    }
    
    # Start background task
    background_tasks.add_task(
        process_summarization,
        job_id,
        request.folder_path,
        request.recursive
    )
    
    return {
        "job_id": job_id,
        "status": "started",
        "message": f"Summarization started for folder: {request.folder_path}"
    }

@app.get("/api/summarize/{job_id}", response_model=SummarizeResponse)
async def get_summarization_status(job_id: str):
    """
    Get the status of a summarization task
    """
    task = task_status.get(job_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found")
    
    response = {
        "job_id": job_id,
        "status": task["status"],
        "message": task["message"]
    }
    
    if task["status"] == "completed":
        response["summary"] = task["result"]
    
    return response

def process_summarization(job_id: str, folder_path: str, recursive: bool):
    """
    Background task to process OneDrive folder summarization
    """
    try:
        result = process_onedrive_folder(folder_path, recursive)
        task_status[job_id].update({
            "status": "completed",
            "message": "Summarization completed successfully",
            "result": result,
            "end_time": datetime.utcnow()
        })
    except Exception as e:
        task_status[job_id].update({
            "status": "failed",
            "message": str(e),
            "end_time": datetime.utcnow()
        })

def list_onedrive_files(access_token: str, folder_path: str, recursive: bool = True) -> List[Dict[str, Any]]:
    """List files in a OneDrive folder using Microsoft Graph API."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # URL encode the folder path
    encoded_path = quote(folder_path.strip('/'))
    # Construct the API URL
    url = f"https://graph.microsoft.com/v1.0/users/drive/root:/{encoded_path}:/children"
    
    # Add query parameters
    params = {
        "$select": "id,name,file,folder,parentReference",
        "$top": 1000  # Maximum number of items to return per page
    }
    
    all_items = []
    
    try:
        while url:
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            data = response.json()
            
            # Get the items from current page
            items = data.get('value', [])
            all_items.extend(items)
            
            # Check if there are more pages
            url = data.get('@odata.nextLink')
            
            # Clear params after first request as they're included in nextLink
            params = {}
            
            
        print(f"Found {len(all_items)} items in folder: {folder_path}")
        
        # If recursive, get items from subfolders
        if recursive:
            # Filter only folders
            folders = [item for item in all_items if 'folder' in item]
            print(f"Found {len(folders)} subfolders to process")
            
            for folder in folders:
                folder_name = folder['name']
                parent_path = folder.get('parentReference', {}).get('path', '').split('root:')[-1].lstrip('/')
                full_path = f"{parent_path}/{folder_name}" if parent_path else folder_name
                print(f"Processing subfolder: {full_path}")
                
                try:
                    subfolder_items = list_onedrive_files(access_token, full_path, recursive=True)
                    all_items.extend(subfolder_items)
                except Exception as e:
                    print(f"Error processing subfolder {full_path}: {str(e)}")
        
        # Filter out folders and only return files
        files = [item for item in all_items if 'file' in item]
        print(f"Total files found: {len(files)}")
        
        return files
        
    except Exception as e:
        error_msg = f"Error listing OneDrive files in {folder_path}: {str(e)}"
        if hasattr(e, 'response') and e.response is not None:
            error_msg += f" | Status: {e.response.status_code} | Response: {e.response.text}"
        print(error_msg)
        raise HTTPException(
            status_code=500,
            detail=error_msg
        )

@app.post("/validate/onedrive", response_model=OneDriveValidationResponse)
async def validate_onedrive_folder(
    request: OneDriveValidationRequest,
    authorization: str = Header(..., description="Bearer token for OneDrive authentication"),
    fastapi_request: Request = None
):
    """
    Validate a document in OneDrive/SharePoint using application permissions.
    The OneDrive access token should be provided in the Authorization header as 'Bearer <token>'.
    """
    print('hello')
    try:
        # Extract the token from the Authorization header
        print('token extracting started')
        token = authorization.split("Bearer ")[1] if authorization.startswith("Bearer ") else authorization
        print("Successfully extracted access token")
        
        
        # Get the OneDrive user ID from the JWT token
        onedrive_user_id = get_onedrive_user_id(fastapi_request)
        print(f"Extracted OneDrive user ID: {onedrive_user_id}")
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # 1. Get the default drive (OneDrive for Business)
        print("\nGetting default drive...")
        
        drive_response = requests.get(
            "https://graph.microsoft.com/v1.0/me/drive",
            headers=headers
        )
        
        if drive_response.status_code != 200:
            error_msg = f"Error getting default drive: {drive_response.text}"
            print(error_msg)
            raise HTTPException(status_code=400, detail=error_msg)
            
        drive_id = drive_response.json().get('id')
        print(f"Using drive ID: {drive_id}")
        
        # 2. List root folder contents
        print("\nListing root folder contents...")
        items_response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children",
            headers=headers
        )
        
        if items_response.status_code != 200:
            error_msg = f"Error listing folder contents: {items_response.text}"
            print(error_msg)
            
            raise HTTPException(status_code=400, detail=error_msg)
            
        items = items_response.json().get('value', [])
        print(f"Found {len(items)} items in root folder")
        
        # If no file path was provided, return the list of items
        if not request.root_folder_name or request.root_folder_name.strip() == '':
            return {
                "status": "success",
                "message": f"Found {len(items)} items in root folder. Please specify a file path.",
                "items": [{"name": item.get('name'), "type": "folder" if 'folder' in item else 'file'} for item in items]
            }
        
        # 3. Try to access the requested file
        file_path = request.root_folder_name.strip('/')
        file_name = os.path.basename(file_path)  # Extract just the file name
        folder_path = os.path.dirname(file_path)  # Get the directory path
        
        validation_errors = []
        file_id = get_item_id(
            folder_path=folder_path,
            onedrive_user_id=onedrive_user_id,
            access_token=token,
            file_value=file_name,
            validation_errors=validation_errors
        )
        
        if not file_id:
            error_details = [vars(error) for error in validation_errors]
            print(f"Error details: {error_details}")
            raise HTTPException(
                status_code=400,
                detail={
                    "status": "error",
                    "message": "Failed to find file",
                    "errors": error_details
                }
            )
        
        # 4. Get file metadata using the file ID
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
        print(f"\nGetting file metadata from: {url}")
        
        file_response = requests.get(url, headers=headers)
        print(f"File metadata status: {file_response.status_code}")
        
        if file_response.status_code != 200:
            error_msg = f"Error accessing file: {file_response.text}"
            print(error_msg)
            raise HTTPException(status_code=400, detail=error_msg)
            
        file_info = file_response.json()
        print(f"Successfully retrieved file: {file_info.get('name')} (ID: {file_id})")
        
        # 5. Download the file content
        if '@microsoft.graph.downloadUrl' in file_info:
            download_url = file_info['@microsoft.graph.downloadUrl']
            file_content = requests.get(download_url).content
            
            # Process the file content as needed
            # ...
            
            return {
                "status": "success",
                "message": f"Successfully processed file: {file_info.get('name')}",
                "file_info": {
                    "name": file_info.get('name'),
                    "size": file_info.get('size'),
                    "last_modified": file_info.get('lastModifiedDateTime'),
                    "web_url": file_info.get('webUrl')
                }
            }
        else:
            raise HTTPException(status_code=400, detail="Download URL not found in file metadata")
            
    except HTTPException as he:
        raise he
    except Exception as e:
        error_msg = f"Unexpected error: {str(e)}"
        print(error_msg)
        raise HTTPException(status_code=500, detail=error_msg)

def get_item_id(folder_path: str, onedrive_user_id: str, access_token: str, file_value: str,
               validation_errors: List[ValidationErrorDetail]) -> Optional[str]:
    """
    Get the ID of a file in a OneDrive folder.
    
    Args:
        folder_path: Path to the folder in OneDrive
        onedrive_user_id: ID of the OneDrive user
        access_token: OAuth access token for Microsoft Graph API
        file_value: Name of the file to find
        validation_errors: List to append validation errors to
        
    Returns:
        str: The ID of the file if found, None otherwise
    """
    try:
        # URL encode the folder path
        encoded_folder_path = quote(folder_path)
        
        # Construct the API URL
        api_url = f"https://graph.microsoft.com/v1.0/users/{onedrive_user_id}/drive/root:/{encoded_folder_path}:/children"
        print(f"Item ID URL: {api_url}")

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }

        # Make the API request
        response = requests.get(api_url, headers=headers)
        response.raise_for_status()  # Raises an exception for 4XX/5XX responses

        # Parse the response
        response_data = response.json()
        items = response_data.get('value', [])
        
        # Search for the file
        for item in items:
            if item.get('name') == file_value:
                return item.get('id')
        
        # If we get here, the file wasn't found
        parent_folder = folder_path[:folder_path.rfind("/") + 1] if "/" in folder_path else ""
        validation_errors.append(
            ValidationErrorDetail(
                folder_path=parent_folder,
                file_name=file_value,
                error_code="15.8",
                status="Fail",
                message="There should be no unreferenced files in M1, M2, M3, M4 and M5 folders",
                details="Including all subfolders within the m1-m5 folders but excluding 'util' folder and subfolders",
                error_type="Files/Folders"
            )
        )
        return None

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
        if http_err.response.status_code == 404:
            error_msg = f"Folder not found: {folder_path}"
        else:
            error_msg = f"Error accessing OneDrive: {http_err.response.text}"
        print(error_msg)
        
        parent_folder = folder_path[:folder_path.rfind("/") + 1] if "/" in folder_path else ""
        validation_errors.append(
            ValidationErrorDetail(
                folder_path=parent_folder,
                file_name=file_value,
                error_code="15.8",
                status="Error",
                message=error_msg,
                details="Error occurred while trying to access OneDrive folder",
                error_type="API Error"
            )
        )
        return None
        
    except Exception as e:
        error_msg = f"Unexpected error: {str(e)}"
        print(error_msg)
        
        parent_folder = folder_path[:folder_path.rfind("/") + 1] if "/" in folder_path else ""
        validation_errors.append(
            ValidationErrorDetail(
                folder_path=parent_folder,
                file_name=file_value,
                error_code="15.8",
                status="Error",
                message=error_msg,
                details="An unexpected error occurred",
                error_type="System Error"
            )
        )
        return None
    




# Add this new endpoint to api.py (before the health check endpoint)
@app.post("/analyze/pharmaceutical", response_model=PharmaceuticalAnalysisResponse)
async def analyze_pharmaceutical_document_api(
    file: Optional[UploadFile] = File(None),
    document_url: Optional[str] = Form(None),
    document_text: Optional[str] = Form(None),
    request: Request = None
):
    """
    Analyze pharmaceutical documents to extract key entities like manufacturers, 
    products, substances, and regulatory references.
    
    Supports:
    - File upload (Word documents or text files)
    - Document URL
    - Direct text input
    
    Returns extracted entities including:
    - Manufacturers/Companies
    - Products/Devices  
    - Active Substances/Ingredients
    - Regulatory References
    """
    try:
        # Check if request is JSON
        content_type = request.headers.get('content-type', '')
        
        if 'application/json' in content_type:
            # Handle JSON request
            json_data = await request.json()
            document_url = json_data.get('document_url')
            document_text = json_data.get('document_text')
            file = None
        
        temp_dir = tempfile.mkdtemp()
        
        try:
            file_path = None
            
            # Handle file upload
            if file:
                # Validate file type
                file_extension = os.path.splitext(file.filename or 'document')[1].lower()
                if file_extension not in ['.docx', '.doc', '.txt']:
                    raise HTTPException(
                        status_code=400,
                        detail="Unsupported file type. Please upload Word documents (.docx, .doc) or text files (.txt)"
                    )
                
                # Create a secure filename
                safe_filename = f"{uuid.uuid4()}{file_extension}"
                file_path = os.path.join(temp_dir, safe_filename)
                
                # Save the uploaded file
                with open(file_path, "wb") as buffer:
                    shutil.copyfileobj(file.file, buffer)
                
                # Analyze the pharmaceutical document (SYNCHRONOUS call)
                analysis_result = analyze_pharmaceutical_content(file_path)
                
                return PharmaceuticalAnalysisResponse(**analysis_result)
                
            # Handle URL or text input if no file was uploaded
            elif document_url or document_text:
                # Create a temporary file for text content
                file_path = os.path.join(temp_dir, f"{uuid.uuid4()}.txt")
                with open(file_path, "w", encoding="utf-8") as f:
                    if document_url:
                        f.write(f"[URL Content: {document_url}]\n\n")
                        # Note: In production, you would download the URL content here
                    if document_text:
                        f.write(document_text)
                
                # Analyze the pharmaceutical document (SYNCHRONOUS call)
                analysis_result = analyze_pharmaceutical_content(file_path)
                
                return PharmaceuticalAnalysisResponse(**analysis_result)
                
            else:
                raise HTTPException(
                    status_code=400,
                    detail="Either 'file', 'document_url', or 'document_text' must be provided"
                )
                
        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"Error during pharmaceutical analysis: {str(e)}"
            )
        finally:
            # Clean up temporary directory
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing pharmaceutical analysis request: {str(e)}"
        )

# Add this endpoint for batch processing multiple documents
@app.post("/analyze/pharmaceutical/batch")
async def analyze_pharmaceutical_batch_api(
    files: List[UploadFile] = File(...)
):
    """
    Analyze multiple pharmaceutical documents in batch.
    
    Returns analysis results for each document with extracted entities.
    """
    job_id = str(uuid.uuid4())
    temp_dir = tempfile.mkdtemp()
    
    try:
        results = []
        
        for file in files:
            try:
                # Validate file type
                file_extension = os.path.splitext(file.filename or 'document')[1].lower()
                if file_extension not in ['.docx', '.doc', '.txt']:
                    results.append({
                        "file": file.filename,
                        "status": "error",
                        "error": f"Unsupported file type: {file_extension}"
                    })
                    continue
                
                # Create a secure filename
                safe_filename = f"{uuid.uuid4()}{file_extension}"
                file_path = os.path.join(temp_dir, safe_filename)
                
                # Save the uploaded file
                with open(file_path, "wb") as buffer:
                    shutil.copyfileobj(file.file, buffer)
                
                # Analyze the pharmaceutical document (SYNCHRONOUS call)
                analysis_result = analyze_pharmaceutical_content(file_path)
                analysis_result["original_filename"] = file.filename
                results.append(analysis_result)
                
            except Exception as e:
                results.append({
                    "file": file.filename,
                    "status": "error",
                    "error": str(e)
                })
        
        return {
            "job_id": job_id,
            "status": "completed",
            "processed_files": len(files),
            "results": results
        }
        
    finally:
        # Clean up temporary directory
        shutil.rmtree(temp_dir, ignore_errors=True)

# Add this endpoint to get pharmaceutical entity statistics
@app.get("/analyze/pharmaceutical/stats")
async def get_pharmaceutical_stats_api():
    """
    Get statistics about pharmaceutical analysis capabilities.
    """
    substance_patterns = [
        "ine$", "ate$", "ide$", "one$", "ol$", "ene$", "ium$", "gen$", "xide$", "acid$",
        "amine$", "zole$", "caine$", "sulfa", "cycline$", "mycin$", "dipine$", "pril$",
        "sartan$", "lol$", "prazole$", "tidine$", "zosin$", "triptan$", "oxetine$"
    ]
    
    return {
        "supported_entity_types": [
            "manufacturers",
            "products", 
            "substances",
            "references"
        ],
        "supported_file_types": [".docx", ".doc", ".txt"],
        "substance_patterns": substance_patterns,
        "analysis_capabilities": [
            "Manufacturer identification",
            "Product name extraction", 
            "Active substance detection",
            "Regulatory reference finding"
        ]
    }



@app.get("/healthy")
def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "timestamp": datetime.utcnow().isoformat(),
        "version": "1.0.0"
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)








    