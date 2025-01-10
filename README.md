# Monitoring Print Events in Windows

This repository contains a Python script that accesses and lists print-related events in the **Microsoft-Windows-PrintService/Operational** log using the `win32evtlog` API.

## Objective
Monitor print events recorded in the Windows Event Viewer, specifically filtering events with ID **307**, which indicate the completion of a print job.

## Features
- Access to the **Microsoft-Windows-PrintService/Operational** event log.
- Counting all events recorded in the log.
- Filtering for events with ID **307**.
- Detailed display of the first filtered events, including:
  - Event ID
  - Source
  - Date and time of the event
  - Associated message

## Dependencies
This project uses the `pywin32` library to access system logs in the Windows operating system.

### Installing the Library
Run the following command to install the library:

```bash
pip install pywin32
```

## Usage
Run the script to list and filter print-related events:

```bash
python monitor_print_events.py
```

### Main Code
```python
import win32evtlog

# Function to list and count print service events
def list_printservice_events(log_name="Microsoft-Windows-PrintService/Operational", event_id=307):
    server = None  # Local machine
    try:
        # Open the operational log
        handle = win32evtlog.OpenEventLog(server, log_name)
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
        
        filtered_events = []
        total_events = 0

        while True:
            events = win32evtlog.ReadEventLog(handle, flags, 0)
            if not events:
                break

            for event in events:
                total_events += 1  # Counting all events
                if event.EventID == event_id:
                    filtered_events.append(event)

        print(f"Total events in log '{log_name}': {total_events}")
        print(f"Events with ID {event_id}: {len(filtered_events)}")

        if filtered_events:
            print("\nDetails of the first 10 filtered events:\n")
            for i, event in enumerate(filtered_events[:10], start=1):
                print(f"Event {i}:")
                print(f"  Event ID: {event.EventID}")
                print(f"  Source: {event.SourceName}")
                print(f"  Time: {event.TimeGenerated}")
                print(f"  Message: {event.StringInserts}")
                print("-" * 50)
        else:
            print(f"No events with ID {event_id} found in log '{log_name}'.")

        win32evtlog.CloseEventLog(handle)

    except Exception as e:
        print(f"Error accessing log '{log_name}': {e}")


# Function call
list_printservice_events()
```

## How It Works
1. **Accessing the Specific Log**: The **Microsoft-Windows-PrintService/Operational** log is directly opened by its name.
2. **Filtering by Event ID**: Only events with ID **307** (indicating a completed print job) are counted and displayed.
3. **Detailed Display**: Detailed information about the first filtered events is shown in the console.

## Requirements
- Windows 10 or later.
- Python 3.x.
- Access to the Windows Event Viewer.

## Testing
1. Run the script as **administrator** to ensure access to the logs.
2. Ensure the **Microsoft-Windows-PrintService/Operational** log is enabled in the Event Viewer:
   - Open the Event Viewer.
   - Navigate to **Applications and Services Logs > Microsoft > Windows > PrintService > Operational**.
   - Enable the log if it is disabled.

## Contributions
Contributions are welcome! Feel free to open an issue or submit a pull request.

