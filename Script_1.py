import mailbox
import os

# Path to your MBOX file (you uploaded it)
mbox_path = "/mnt/data/your_file.mbox"

# Output folder
output_dir = "attachments"
os.makedirs(output_dir, exist_ok=True)

# Open MBOX
mbox = mailbox.mbox(mbox_path)

for i, message in enumerate(mbox):
    for part in message.walk():
        content_disposition = part.get("Content-Disposition")

        if content_disposition and "attachment" in content_disposition:
            filename = part.get_filename()

            if filename:
                filepath = os.path.join(output_dir, filename)

                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))

                print(f"Saved: {filename}")

print("Done extracting attachments.")