import openpyxl
import pandas as pd
from collections import Counter


def get_most_frequent_country_from_cities(cities):
    """
    Load an Excel file, retrieve country names for a list of cities,
    and return the country with the maximum frequency.

    Parameters:
    - cities: list of str, city names to look up

    Returns:
    - The country with the highest frequency from the list of cities
    """
    # Hardcoded path to the Excel file
    path = r"C:\Users\Z19661\PycharmProjects\LegacyData_\CODES\database\countrydata.xlsx"

    # Load the Excel file
    workbook = pd.read_excel(path)



    # Ensure input is a list (for single city, convert to list)
    if isinstance(cities, str):
        cities = [cities]  # Convert single city name to a list

    countries = []
    not_found = []  # Keep track of cities that are not found

    for city in cities:
        # Get the country name for the city (case-insensitive search)
        country = workbook.loc[workbook['City'].str.lower() == city.lower(), 'Country']

        # If the city is found, append the country name; otherwise, append None
        if not country.empty:
            countries.append(country.iloc[0])
        else:
            countries.append(None)
            not_found.append(city)  # Track the city not found

    # Print cities that were not found in the data
    if not_found:
        print(f"Cities not found: {', '.join(not_found)}")

    # Remove None values (cities that were not found)
    countries = [country for country in countries if country is not None]

    if not countries:
        return None  # If no valid countries are found

    # Count the frequency of each country
    country_counts = Counter(countries)

    # Get the country with the highest frequency
    most_frequent_country = country_counts.most_common(1)[0][0]
    return most_frequent_country


# List of cities to search for
# cities = [
#     "Ho Chi Minh",
#     "Hanoi",
#     "Flanders",
#     "Wallonia",
#     "Brussels",
#     "São Paulo",
#     "São Paulo",
#     "Porto Alegre",
#     "Santiago",
#     "Concepcion",
#     "Seoul",
#     "Mexico",
#     "Monterrey",
#     "Guadalajara",
#     "Esmeralda 1 - Buenos Aires"
# ]
#
# # Print the most frequent country from the list of cities
# print(get_most_frequent_country_from_cities(cities))
