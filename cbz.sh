#!/usr/bin/env bash
# ----------------------------------------------------------------------
# --
# -- Created by PonyLucky.
# --
# -- version: 0.1.
# --
#
# -- Description
# The function of this script is to convert an image container into a
# CBZ file (used for comics or mangas, readable on Ebooks and any
# devices with the good application).
#
# -- Prerequises
# You need the package `zip` installed.
#
# -- How does this script work
# This script ZIPs all sub-directories containing only images in the
# current directory and changes their extension to .cbz.
#
# -- Notes
# If you take this script into another one or a project, credits will
# be appreciated. Though it is not an obligation.
# ----------------------------------------------------------------------

# Check that zip is installed
if ! command -v zip &> /dev/null; then
    echo "Error: 'zip' is not installed. Please install it first."
    exit 1
fi

# Ask if we delete folders after conversion
read -p "Do you want to delete the folders after conversion [Y/N]: " is_del
echo
is_del="${is_del:-N}"
if [[ "${is_del^^}" != "Y" && "${is_del^^}" != "N" ]]; then
    echo "Error: invalid input '${is_del}'. Please enter Y or N."
    exit 1
fi

echo "Converted:"

converted=0

for dir in */; do
    dir="${dir%/}"
    [ -d "$dir" ] || continue

    # Count total files vs image files in the directory (recursively)
    total=$(find "$dir" -type f | wc -l)
    images=$(find "$dir" -type f \( \
        -iname "*.jpg"  -o -iname "*.jpeg" -o \
        -iname "*.png"  -o -iname "*.gif"  -o \
        -iname "*.webp" -o -iname "*.bmp"  -o \
        -iname "*.tiff" -o -iname "*.tif"  -o \
        -iname "*.avif" \
    \) | wc -l)

    # Skip empty directories or those with non-image files
    [ "$total" -eq 0 ] && continue
    [ "$total" -ne "$images" ] && continue

    output="${dir}.cbz"
    (cd "$dir" && zip -r "../${output}" . -x ".*") > /dev/null

    [[ "${is_del^^}" == "Y" ]] && rm -rf "$dir"

    echo "-- ${dir}.cbz"
    ((converted++))
done

echo
if [ "$converted" -eq 0 ]; then
    echo "No folders with only images found."
else
    echo "Finished"
fi
