cd ../
virtualenv venv
call ./venv/Scripts/activate
pip install -r ./batch/requirements.txt
call ./venv/Scripts/deactivate
pause