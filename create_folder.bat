@echo off
if not exist "dest_folder" (
    mkdir dest_folder
    echo Created dest_folder
) else (
    echo dest_folder already exists
)

if not exist "images" (
    mkdir images
    echo Created images
) else (
    echo images already exists
)
