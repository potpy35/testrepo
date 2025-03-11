import requests
import sys
import os
from openpyxl import Workbook, load_workbook

# TMDb API Key
API_KEY = "4fb3b577e9c675bd408d6d22b4e8ed54"
BASE_URL = "https://api.themoviedb.org/3"
EXCEL_FILENAME = os.path.expanduser("~/Library/Mobile Documents/iCloud~com~omz-software~Pythonista3/Documents/Shortcut Automations/Watchlist/Watchlist.xlsx")  # File location on iOS (Pythonista)

def get_movie_data(title):
    """Fetches movie or TV show metadata from TMDb API."""
    
    search_url = f"{BASE_URL}/search/multi"
    params = {
        "api_key": API_KEY,
        "query": title,
        "include_adult": False
    }
    
    response = requests.get(search_url, params=params)
    if response.status_code != 200:
        print("Error fetching data from API")
        return None
    
    results = response.json().get("results", [])
    if not results:
        print("No results found.")
        return None

    # Get first result (best match)
    result = results[0]
    media_type = result.get("media_type", "movie")  # Could be 'movie' or 'tv'
    
    # Fetch detailed metadata
    if media_type == "movie":
        details_url = f"{BASE_URL}/movie/{result['id']}"
    else:
        details_url = f"{BASE_URL}/tv/{result['id']}"

    details_params = {"api_key": API_KEY}
    details_response = requests.get(details_url, params=details_params)
    
    if details_response.status_code != 200:
        print("Error fetching details.")
        return None
    
    details = details_response.json()

    # Extract common data
    title = details.get("title") or details.get("name", "N/A")
    genre = ", ".join([g["name"] for g in details.get("genres", [])]) or "Unknown"
    synopsis = details.get("overview", "No synopsis available.")

    # Fetch streaming information (via JustWatch integration in TMDb)
    watch_url = f"{BASE_URL}/{media_type}/{result['id']}/watch/providers"
    watch_response = requests.get(watch_url, params={"api_key": API_KEY})
    streaming_platforms = "Not Available"
    
    if watch_response.status_code == 200:
        watch_data = watch_response.json().get("results", {}).get("US", {}).get("flatrate", [])
        if watch_data:
            streaming_platforms = ", ".join([provider["provider_name"] for provider in watch_data])

    # Handle duration
    if media_type == "movie":
        duration = f"{details.get('runtime', 'Unknown')} minutes"
    else:
        seasons = details.get("number_of_seasons", "Unknown")
        avg_episode_time = details.get("episode_run_time", [0])
        avg_duration = f"{sum(avg_episode_time) // len(avg_episode_time)} minutes" if avg_episode_time else "Unknown"
        duration = f"{seasons} season(s), ~{avg_duration} per episode"

    return ["", title, media_type.capitalize(), genre, streaming_platforms, synopsis, duration]

def append_to_excel(data):
    """Appends movie or TV show metadata to an existing Excel file, or creates one if it doesn't exist."""
    
    headers = ["", "Title", "Type", "Genre", "Streaming Platforms", "Synopsis", "Duration"]
    
    if os.path.exists(EXCEL_FILENAME):
        # Load existing workbook
        workbook = load_workbook(EXCEL_FILENAME)
        sheet = workbook.active
    else:
        print("No watchlist file found.")

    # Append new data
    sheet.append(data)

    # Save workbook
    workbook.save(EXCEL_FILENAME)
    print(f"Data successfully saved to {EXCEL_FILENAME}")

if __name__ == "__main__":
    # Get input from Apple Shortcuts via Pythonista
    if len(sys.argv) > 1:
        user_input = sys.argv[1]
    else:
        print("Error: No input received from Shortcuts.")
        sys.exit(1)

    metadata = get_movie_data(user_input)
    
    if metadata:
        append_to_excel(metadata)
