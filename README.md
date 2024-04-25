# MongoDB to Excel Exporter

## Overview
This Python program connects to a specified MongoDB on a remote server via SSH tunneling and exports a specified collection to an Excel file (.xlsx).

## Requirements
- Python 3.x
- Pandas
- PyMongo
- SSHTunnel
- XlsxWriter

## Input
This program does not require any input at runtime. However, you need to modify the following parameters in the script:
- SSH_HOST: Hostname or IP address of the SSH server.
- SSH_USER: Username for SSH authentication.
- SSH_PASSWORD: Password for SSH authentication.
- MONGO_HOST: Hostname or IP address of the MongoDB server.
- DB: Name of the MongoDB database.
- COLLECTION: Name of the collection to export to Excel.

## Output
- The program generates an Excel file named `output.xlsx`, containing the data from the specified MongoDB collection.

## Usage
1. Modify the SSH and MongoDB parameters in the script according to your setup.
2. Ensure you are connected to the University network for the SSH connection to work, or use eduVPN.
3. Run the Python script. If prompted, allow access through your firewall.
4. If the program times out at any point, please try running it again.

## Note
- Make sure you have necessary permissions to access the MongoDB database and SSH server.
- Ensure that the MongoDB server is running and accessible through the SSH tunnel.

## Author
[Malek Miled] - 25.04.2024
