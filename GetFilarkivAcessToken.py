from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection


def GetFilarkivToken(orchestrator_connection: OrchestratorConnection):
    import os
    from datetime import datetime, timedelta
    import pytz
    import requests

    
    FilarkivTokenTimestamp = orchestrator_connection.get_constant("FilarkivTokenTimestamp1").value
    Filarkiv_access= orchestrator_connection.get_credential("FilarkivAccessToken1")
    Filarkiv_access_token = Filarkiv_access.password
    Filarkiv_URL = Filarkiv_access.username


    # Define Danish timezone
    danish_timezone = pytz.timezone("Europe/Copenhagen")

    # Parse the old timestamp to a datetime object
    old_time = datetime.strptime(FilarkivTokenTimestamp.strip(), "%d-%m-%Y %H:%M:%S")
    old_time = danish_timezone.localize(old_time)  # Localize to Danish timezone
    print('Old timestamp: ' + old_time.strftime("%d-%m-%Y %H:%M:%S"))

    # Get the current timestamp in Danish timezone
    current_time = datetime.now(danish_timezone)
    print('current timestamp: '+current_time.strftime("%d-%m-%Y %H:%M:%S"))
    str_current_time = current_time.strftime("%d-%m-%Y %H:%M:%S")

    # Calculate the difference between the two timestamps
    time_difference = current_time - old_time
    print(time_difference)

    # Check if the difference is over 1 hour and 30 minutes
    GetNewTimeStamp = time_difference > timedelta(minutes=30)

    # Output for the boolean
    print("GetNewTimeStamp:", GetNewTimeStamp)

    # Example of using it in an if-statement
    if GetNewTimeStamp:
        print("The difference is over 30 minutes. Fetch a new timestamp!")
        # Replace these values with your actual keys
        client_id = 'fa_de_aarhus_job_user'
        client_secret = 'l661C4TUGI4Xqu1k5qol9KqoJ-BHOYwXSfXx1BZZFnWF0YB!LOil584jwcTQgLZgjf3VvIsh8H6Njr20'
        scope = 'fa_de_api:normal'
        grant_type = 'client_credentials'

        # Data to be sent in the POST request
        keys = {
            'client_secret': client_secret,
            'client_id': client_id,
            'scope': scope,
            'grant_type': grant_type,  # Specify the grant type you're using
        }

        # Sending POST request to get the access token
        response = requests.post(Filarkiv_URL, data=keys)

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            Filarkiv_access_token = response.json().get('access_token')
            print("Access token granted")
            orchestrator_connection.update_credential("FilarkivAccessToken1",Filarkiv_URL,Filarkiv_access_token)
            orchestrator_connection.update_constant("FilarkivTokenTimestamp1",str_current_time)
            return Filarkiv_access_token
        else:
            print("Failed to get the access token")

    else:
        print("No need to fetch a new timestamp - using old timestamp")
        return Filarkiv_access_token