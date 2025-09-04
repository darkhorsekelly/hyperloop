# Screenshot Management System - Google Apps Script

This script automates the process of capturing, resizing, and organizing screenshots within a Google Sheet, designed for efficient data entry workflows.

## Key Decisions & Refactoring Analysis

The original script was functional but had opportunities for improvement in terms of robustness, maintainability, and performance. The following key decisions were made during the refactoring process:

### 1. Modularity and Code Structure
- **Decision**: Break down large functions into smaller, single-purpose helper functions.
- **Rationale**: The original `processImageAtPosition` handled finding, clearing, inserting, and resizing images. This was split into `processSingleImageMapping`, `clearImagesAtCell`, and `resizeReportImage`. This improves readability, makes the code easier to test, and allows for reuse of functions like `clearAllImagesFromSheet`.
- **Impact**: The code is now more organized and easier to follow. Each function has a clear responsibility.

### 2. Comprehensive Error Handling
- **Decision**: Implement a centralized `handleError` function and wrap all user-facing and background operations in `try...catch` blocks.
- **Rationale**: The original script lacked robust error handling, meaning a single failure (e.g., a misconfigured cell name in `CONFIG`) could halt execution without clear feedback. The new system ensures that errors are caught, logged appropriately, and displayed to the user in a non-blocking way.
- **Impact**: The script is significantly more resilient. Users are informed of issues, and background triggers won't fail silently.

### 3. Performance Optimization
- **Decision**: Reduce redundant API calls, especially `getImages()`.
- **Rationale**: Calling `sheet.getImages()` repeatedly within a loop is inefficient. The refactored `processAllScreenshots` now fetches all images from the worker sheet *once* and passes the array down to the processing functions.
- **Impact**: The script runs faster, especially when processing a large number of images, providing a better user experience and reducing the risk of hitting Google Apps Script execution time limits.

### 4. Bug Fixes and Feature Enhancements
- **Critical Bug Fix (Triggers)**: The original script set up a time-based trigger for a function named `autoProcessImages` that did not exist. This was fixed by creating the `autoProcessImages` function and ensuring it correctly calls the main processing logic while suppressing UI elements (which are not allowed in triggers).
- **Functional Image Preview**: The image preview dialog was updated to display the actual image using a Base64 data URI, making it a genuinely useful feature.
- **Refined User Feedback**: Dialogs and prompts were simplified for clarity.

### 5. In-Script Testing Framework
- **Decision**: Add a lightweight, self-contained testing framework (`runTests` function).
- **Rationale**: Since Apps Script lacks a traditional testing environment, creating an in-script test suite is crucial for ensuring reliability. It allows developers to verify core functionality (like image processing) safely using temporary sheets, preventing regressions when making future changes.
- **Impact**: The script's long-term maintainability is vastly improved. New features or fixes can be validated with a single click, increasing confidence in deployments.

## How to Use
1.  **Initial Setup**: Open the script editor (`Extensions > Apps Script`) and run the `initialSetup` function once. This will create the custom menu and set up the necessary triggers.
2.  **Daily Use**:
    *   Paste screenshots into the green cells in the `WorkerDataEntryForm` sheet.
    *   Use the `📸 Screenshots` menu to process them, clear images, or use "Quick Paste Mode" to speed up entry.
3.  **Running Tests**: From the `📸 Screenshots` menu, select `Run Tests`. The script will create temporary sheets, run its verification logic, and then clean up after itself. Check the logs (`View > Logs`) for detailed results.

This refactoring effort has transformed the script from a simple utility into a robust and maintainable tool designed for long-term use.
