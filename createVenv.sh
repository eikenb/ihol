#!/bin/bash

requirementPath=requirements.txt

if [ $# -lt 1 ]; then
    echo " Usage: $0 <venvPath>"
    echo
    exit 1
fi


venvPath=$1;shift

if [ -e $venvPath ];    then
    echo "ERROR: $venvPath path already exists ! Nothing will be done."
    exit 1
fi

if [ ! -f $requirementPath ]; then
    echo "ERROR: $requirementPath : file not found. This script need the file that contain the list of package need to setup the venv."
    exit 1
fi

function runcmd {
    echo "> $@"
    eval "$@"
}

set -e


runcmd mkdir -p $venvPath
runcmd python3 -m venv $venvPath
runcmd source $venvPath/bin/activate
runcmd pip install --upgrade pip
runcmd pip install -r $requirementPath
echo
echo "Virtual env setup succeed. Activate it with \"source $(realpath $venvPath)/bin/activate\""




