# summarize_onedrive.py
from typing import Dict, List, Any
import os
from pathlib import Path

def process_onedrive_folder(folder_path: str, recursive: bool = False) -> Dict[str, Any]:
    """
    Process a OneDrive folder and return a summary of its contents.
    
    Args:
        folder_path (str): Path to the OneDrive folder
        recursive (bool): Whether to process subdirectories
        
    Returns:
        Dict[str, Any]: Summary of the folder contents
    """
    try:
        path = Path(folder_path)
        
        # Check if path exists
        if not path.exists():
            return {
                "status": "error",
                "message": f"Path does not exist: {folder_path}",
                "folder": str(folder_path)
            }
            
        if not path.is_dir():
            return {
                "status": "error",
                "message": f"Path is not a directory: {folder_path}",
                "folder": str(folder_path)
            }
            
        # Initialize summary
        summary = {
            "folder": str(folder_path),
            "file_count": 0,
            "file_types": {},
            "total_size": 0,  # in bytes
            "files": []
        }
        
        # Process files
        for item in path.iterdir():
            if item.is_file():
                file_info = {
                    "name": item.name,
                    "path": str(item),
                    "size": item.stat().st_size,
                    "type": item.suffix.lower() or "unknown",
                    "last_modified": item.stat().st_mtime
                }
                summary["files"].append(file_info)
                summary["file_count"] += 1
                summary["total_size"] += file_info["size"]
                
                # Update file type count
                file_type = file_info["type"]
                summary["file_types"][file_type] = summary["file_types"].get(file_type, 0) + 1
                
            elif item.is_dir() and recursive:
                # Process subdirectories recursively
                sub_summary = process_onedrive_folder(str(item), recursive)
                if sub_summary["status"] == "success":
                    summary["file_count"] += sub_summary["file_count"]
                    summary["total_size"] += sub_summary["total_size"]
                    # Merge file type counts
                    for ftype, count in sub_summary["file_types"].items():
                        summary["file_types"][ftype] = summary["file_types"].get(ftype, 0) + count
        
        # Convert size to MB for readability
        summary["total_size_mb"] = round(summary["total_size"] / (1024 * 1024), 2)
        summary["status"] = "success"
        return summary
        
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
            "folder": str(folder_path)
        }

# Example usage
if __name__ == "__main__":
    # Test with the current directory
    result = process_onedrive_folder(".", recursive=True)
    print("Folder summary:")
    print(f"Total files: {result['file_count']}")
    print(f"Total size: {result['total_size_mb']} MB")
    print("File types:", result["file_types"])