import pandas as pd
import random
import datetime

from typing import Dict, List, Tuple, Optional

# Read the input Excel file into a DataFrame
file_path = "signees.xlsx"
df = pd.read_excel(file_path)

# Get the unique values from the 'GROUP' column in the DataFrame
distinct_groups = df['GROUP'].unique()

# Initialize an empty dictionary to hold lists for each group
groups_dict = {}

# Create an empty list for each distinct group
for group in distinct_groups:
    groups_dict[group] = []

# Populate each group's list with corresponding data from the DataFrame
for group in distinct_groups:
    # Extract entries for the current group and convert to a list of dictionaries
    group_entries = df[df['GROUP'] == group].to_dict('records')
    
    # Shuffle the entries within the group
    random.shuffle(group_entries)
    
    # Write the shuffled list to the group in the dictionary
    groups_dict[group] = group_entries

# Function to pop elements from the longest and second longest lists in the dictionary
def pop_from_longest_and_second_longest_dict(group_dict: Dict[str, List[str]]) -> Tuple[Optional[str], Optional[str]]:
    # Check if the dictionary is empty
    if not group_dict:
        return None, None
    
    # Check if the dictionary has only one group
    elif len(group_dict) == 1:
        # Get the only key in the dictionary
        key = next(iter(group_dict))
        # Pop and return the last element from the only list, or None if the list is empty
        return group_dict[key].pop() if group_dict[key] else None, None
    
    # Sort the dictionary items by the length of their lists in descending order
    sorted_items = sorted(group_dict.items(), key=lambda item: len(item[1]), reverse=True)
    
    # Pop and return the last element from the longest list
    longest_pop = sorted_items[0][1].pop() if sorted_items[0][1] else None
    
    # Pop and return the last element from the second longest list
    second_longest_pop = sorted_items[1][1].pop() if sorted_items[1][1] else None
    
    return (longest_pop, second_longest_pop)
    # ...
    # ...

pairs = []
# Continue pairing members while there are still members left in any group
while any(groups_dict.values()):
    # Pop a pair of members from the longest and second longest lists
    pair = pop_from_longest_and_second_longest_dict(groups_dict)
    
    # Check if the pair is not (None, None)
    if pair != (None, None):
        # Append the valid pair to the pairs list
        pairs.append(pair)

# Iterate over each pair of peers by index
for i, pair in enumerate(pairs):
    # Check if the first peer in the pair is None
    if pair[0] is None:
        # Replace the first peer with a placeholder dictionary if it is None
        pairs[i] = ({'Mail': 'N/A', 'GROUP': 'N/A', 'First name': 'N/A'}, pair[1])
    # Check if the second peer in the pair is None
    elif pair[1] is None:
        # Replace the second peer with a placeholder dictionary if it is None
        pairs[i] = (pair[0], {'Mail': 'N/A', 'GROUP': 'N/A', 'First name': 'N/A'})

# Initialize a dictionary to store the peer information
peer_output_dict = {
    'Peer1 Email': [],
    'Peer1 TG': [],
    'Peer1 First Name': [],
    'Peer2 Email': [],
    'Peer2 TG': [],
    'Peer2 First Name': []
}

# Iterate over each pair of peers
for pair in pairs:
    peer1 = pair[0]  # First peer in the pair
    peer2 = pair[1]  # Second peer in the pair

    # Append the information for Peer1 to the dictionary, handling None values
    peer_output_dict['Peer1 Email'].append(peer1['Mail'] if peer1 is not None else None)
    peer_output_dict['Peer1 TG'].append(peer1['GROUP'] if peer1 is not None else None)
    peer_output_dict['Peer1 First Name'].append(peer1['First name'] if peer1 is not None else None)

    # Append the information for Peer2 to the dictionary, handling None values
    peer_output_dict['Peer2 Email'].append(peer2['Mail'] if peer2 is not None else None)
    peer_output_dict['Peer2 TG'].append(peer2['GROUP'] if peer2 is not None else None)
    peer_output_dict['Peer2 First Name'].append(peer2['First name'] if peer2 is not None else None)

# Convert the dictionary to a DataFrame
peer_output_df = pd.DataFrame(peer_output_dict)

# Get the current timestamp in the format YYYYMMDD_HHMMSS
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# Create a filename using the timestamp
filename = f'peers_data_{timestamp}.xlsx'

# Save the DataFrame to an Excel file
peer_output_df.to_excel(filename, index=False)
