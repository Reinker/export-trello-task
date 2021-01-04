#!/bin/bash
function setVirtualEnv() {
	source venv/bin/activate
	pip install pylint
	pip install Jedi
	pip install numpy
	pip install openpyxl
}

if virtualenv -p python3.7 venv > /dev/null 2>&1;then
	setVirtualEnv
else 
	pip install virtualenv
	virtualenv -p python3.7 venv
	setVirtualEnv
fi
