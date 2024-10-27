import streamlit as st
import pandas as pd
import os

# Function to check if the Excel file exists
def excel_file_exists():
    return os.path.exists('events.xlsx')

# Function to create the Excel file with dummy data
def create_excel_file():
    dummy_data = {
        'Event': ['Event 1'],
        'Name': ['Dummy Event'],
        'Venue': ['Dummy Venue'],
        'Date': ['2024-01-01'],
        'Time': ['10:00 AM'],
        'Guest Speaker': ['Dummy Speaker'],
        'Description': ['This is a dummy event description.']
    }
    df = pd.DataFrame(dummy_data)
    df.to_excel('events.xlsx', index=False)

# Function to load data from the Excel file
def load_data():
    if not excel_file_exists():
        create_excel_file()
    return pd.read_excel('events.xlsx', dtype={'Date': 'datetime64[ns]'})  # Explicitly specify dtype

# Function to create new event
def create_new_event(event_name, name, venue, date, time, guest_speaker, description):
    df = load_data()
    new_row = {'Event': event_name, 'Name': name, 'Venue': venue, 'Date': date, 'Time': time,
               'Guest Speaker': guest_speaker, 'Description': description}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)  # Use pd.concat instead of append
    df.to_excel('events.xlsx', index=False)
    st.success('New event added successfully.')
    return df

# Function to update event details
def update_event_details(old_event_name, event_name, name, venue, date, time, guest_speaker, description):
    df = load_data()
    df.loc[df['Event'] == old_event_name, ['Event', 'Name', 'Venue', 'Date', 'Time', 'Guest Speaker', 'Description']] = \
        event_name, name, venue, date, time, guest_speaker, description
    df.to_excel('events.xlsx', index=False)
    st.success('Event details updated successfully.')
    return df

# Main Streamlit app
st.title('Event Program Generator')

# Sidebar with options
option = st.sidebar.selectbox('Select an Option', ['Select an Option', 'List All Events', 'Create New Event'] + load_data()['Event'].tolist())
if option == 'Create New Event':
    with st.form('create_new_event_form'):
        event_name = st.text_input('Event Name')
        name = st.text_input('Name')
        venue = st.text_input('Venue')
        date = st.date_input('Date')
        time = st.time_input('Time')
        guest_speaker = st.text_input('Guest Speaker')
        description = st.text_area('Description')

        submit_button = st.form_submit_button(label='Add Event')
        
        if submit_button:
            create_new_event(event_name, name, venue, date, time, guest_speaker, description)
            st.experimental_rerun()  # Reload the page to update dropdown options
            # Clear input fields
            st.session_state.create_new_event_form = {
                'event_name': '',
                'name': '',
                'venue': '',
                'date': '',
                'time': '',
                'guest_speaker': '',
                'description': ''
            }
elif option == 'List All Events':
    st.subheader('All Events')
    all_events = load_data()
    st.write(all_events)
elif option != 'Select an Option':
    if option != 'Create New Event':
        with st.form('update_event_form'):
            existing_event = load_data().loc[load_data()['Event'] == option].iloc[0]
            event_name = st.text_input('Event Name', value=existing_event['Event'], key='event_name')
            name = st.text_input('Name', value=existing_event['Name'], key='name')
            venue = st.text_input('Venue', value=existing_event['Venue'], key='venue')
            date = st.date_input('Date', value=pd.to_datetime(existing_event['Date']).date(), key='date')
            time = st.time_input('Time', value=pd.to_datetime(existing_event['Time']).time(), key='time')
            guest_speaker = st.text_input('Guest Speaker', value=existing_event['Guest Speaker'], key='guest_speaker')
            description = st.text_area('Description', value=existing_event['Description'], key='description')

            submit_button = st.form_submit_button(label='Update Event Details')
            
            if submit_button:
                update_event_details(existing_event['Event'], event_name, name, venue, date, time, guest_speaker, description)
                st.experimental_rerun()  # Reload the page to update dropdown options
