# Utilities for MS Access

These utilities help facilitate working collaboratively with MS-Access applications.  They allow the source code of the application to be exported and imported, thus they can be saved to source control so that changes can be tracked, merged, etc.

## Expected locations, etc

 * create a "utilities" folder
 * download these files (or use git submodules, or SVN externals)
 * modify as necessary

## Files and Their Uses

 * ```decompose.vbs``` (and ```decompose.cmd```)
 
Running this script will export most of the important code out of the *.mdb file, storing them in a "Source" folder (with "Forms," "Macros," "Modules," and "Reports" sub-folders).  The *.cmd file is for running the *.vbs script.

 * ```rebuildFromTextFiles.vbs``` (and ```rebuild.cmd```)
 
 Imports files from the Source folder back into the target database.
 
  * ```upsize--remove_dbo_prefix.vbs``` (and ```remove_dbo_prefix.cmd```)
 
 Removes the "dbo_" prefix from all tables after going through the SQL Server upsize wizard.