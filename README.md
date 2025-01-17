# HideTeamsPopup

This little Python script I did made because I'm very disgusted from this little popup which pops out every time I minimize a meeting in Microsoft Teams ©. This only supports new version because classic is deprecated.

## Getting started

### Prerequisites

- Microsoft Windows © (only tested on 10 and higher)
- Python (v3.12.x and higher)
- pip (v24.x and higher)

### Installation

1. Clone the repository:

   ```sh
   git clone https://github.com/n0vedad/HideTeamsPopup.git
   cd HideTeamsPopup
   ```

2. Install the dependencies:

   ```sh
   pip install -r requirements.txt
   ```

#### Development mode

Simply run the script in your chosen IDE or in a terminal with `python HideTeamsPopup.py`. 

#### Production mode

Install `pyinstaller` with pip and execute the following command in the project folder: `pyinstaller --onefile --noconsole --add-data "icons\icon.ico;icons" --icon="icons\icon.ico" HideTeamsPopup.py`. This will give you a executable file. 

### Behaviour

Once executed the script will run in the background, sitting in the systray and watch for a popup appearing. To exit click on the icon in the systray and choose 'Exit'.

## License

This project is licensed under the MIT License. See the [LICENSE](/LICENSE) for details.
