# ClubHouse.io Utilities
Some sample scripts for managing Clubhouse.io API

**to_excel_sample.py**

Searches a Clubhouse.io organization for a particular set of stories and outputs them to an Excel xlsx file with sheets for each Epic.

## Clubhouse https://clubhouse.io/
Clubhouse is a project management platform for software teams that provides the perfect balance of simplicity and structure. 

## Getting started
Assuming a version of [Python 3.6.4 or higher is installed ](https://docs.python.org/3/) and some knowledge of Python's virtual environments. I personally recomend using [JetBrain's PyCharm](https://www.jetbrains.com/pycharm/) to make life so much easier, but it's not everyone's favourite IDE. If using PyCharm, I typically setup the environmemt variable within the configuration or project and use the Run button to run the configuration as needed. The steps below may assist if you chose not to use PyCharm.

### Create and activate a virtual environment
Create a folder on your local drive.
For more instructions see the [official documentation for virtual environments](https://docs.python.org/3/library/venv.html)
```
python3 -m venv venv
.\venv\Scripts\activate
```
### Check out a version of repository
Use the clone or download button above and save to your project folder.

### Install the requirements
Ensuring the virtual environment is active, install the requirements.
```
pip install -r requirements.txt
```
### Environment variables
An environment variable to house your [Clubhouse  API token](https://help.clubhouse.io/hc/en-us/articles/205701199-Clubhouse-API-Tokens) must be created and called CLUBHOUSE_TOKEN.

### Understand the code
Look to the main() method for description of what is going to happen. Be sure to understand what's hapenning before running any code.
