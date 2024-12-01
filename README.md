## Warm and Fuzzies Script

Warm and Fuzzies are warm messages of appreciation or friendliness that people can send to each other.

The basic flow is: 
- The participants are sent a google form to fill in their warm and fuzzies.
- The responses are used (via this script) to generate a beautifully formatted document for each person's warm and fuzzies. 
    - Optionally: The code can pre-generate generic warm and fuzzies using a sheet of contacts, ensuring everybody recieves a warm and fuzzy 🤗.
- Another script is used to send the attachments to their recipients.

## Usage

The instructions for the broader Warm and Fuzzies process can be found here: https://github.com/matthew-lim-matthew-lim/warm-and-fuzzies

We will proceed with the instructions for this step of creating the warm and fuzzies.

- Prepare a `responses.xlsx` file, which will contain form responses for your warm and fuzzies.
- Prepare a `contacts.xlsx` file, which will contain the contact details of the people you will be recieving warm and fuzzies.
- **[Optional: Generic messages for all contacts]** modify `generic_warm_and_fuzzies.py` to change the generic message sent to all recipients. You do not need to run this file.
- Run `process.py`. The warm and fuzzies will all be in the folder `output_warm_and_fuzzies`.