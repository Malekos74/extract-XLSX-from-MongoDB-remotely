"""
    Python program that helps you connect to a specified mongoDB on a remote server.
        
    INPUT:
        
    OUTPUT:
        - xlsx file of the corresponding collection
    
    NB:
        - You need to be connected to the Uni network for the SSH connection to work (Or use eduVPN)
        - If the program times out at any point, please try again.
        
[Malek Miled] 25.04.2024
"""
import pandas as pd
from pymongo import MongoClient
from sshtunnel import SSHTunnelForwarder

# SSH tunnel parameters (Can be changed accordingly)
SSH_HOST = 'HOST'
SSH_USER = 'USERNAME'
SSH_PASSWORD = 'PASSWORD'
SSH_PORT = 22  # Default SSH port

# MongoDB parameters (Can be changed accordingly)
MONGO_HOST = '127.0.0.1'  # This points to localhost because MongoDB will be accessed through the SSH tunnel
MONGO_PORT = 27017
# Specify the needed DB and COLLECTION as you see fit
DB = 'CHP'
COLLECTION = 'characterization'

if __name__ == '__main__':
    try:
        # Connect to MongoDB through the SSH tunnel
        with SSHTunnelForwarder(
                (SSH_HOST, SSH_PORT),
                ssh_username=SSH_USER,
                ssh_password=SSH_PASSWORD,
                remote_bind_address=('127.0.0.1', MONGO_PORT)
        ) as tunnel:
            client = MongoClient(MONGO_HOST, tunnel.local_bind_port)
            db = client[DB]
            collection = db[COLLECTION]

            # Retrieve documents from MongoDB
            documents = collection.find()

            # Convert documents to a list of dictionaries
            data_list = []
            for doc in documents:
                # Convert ObjectId to string
                doc['_id'] = str(doc['_id'])
                data_list.append(doc)

            # Create a DataFrame from the list of dictionaries
            df = pd.DataFrame(data_list)

            # Save DataFrame to Excel
            writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.close()

            print("Data successfully exported to output.xlsx")

    except Exception as e:
        print("An error occurred:", e)
