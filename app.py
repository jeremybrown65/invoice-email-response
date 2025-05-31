import streamlit as st
import pandas as pd
import os
import io
from pdf2image import convert_from_bytes
from PIL import Image

# --- Constants ---
STORE_LIST_PATH = "data/store_list.xlsx"

RESPONSE_TEMPLATES = [
    {"option": "One store", "needs_store": True, "attach_file": False,
     "reply": "This bills to: GL code 170.3010.XXXXX.000.6340.623020.000.0000"},
    {"option": "All stores", "needs_store": False, "attach_file": True,
     "reply": "Please allocate evenly across all stores. List of stores with Region Codes attached."},
    {"option": "Group of Stores", "needs_store": True, "attach_file": True,
     "reply": "Please allocate evenly across the list of stores with Region Codes attached."},
    {"option": "Lab Store", "needs_store": False, "attach_file": False,
     "reply": "This invoice is for the Lab Store. Please bill to: GL code 170.3010.10125.000.6340.623020.000.0000"},
    {"option": "Retail Activations - Dallas", "needs_store": False, "attach_file": False,
     "reply": "This bills to: GL code: 170.3010.15910.6340.632020"},
    {"option": "Retail Activations - Trailer", "needs_store": False, "attach_file": True,
     "reply": "This bills to: GL code: 170.3010.15916.6340.632020"},
    {"option": "Retail Activations - General", "needs_store": False, "attach_file": False,
     "reply": "This bills to: GL code:"},
    {"option": "Scrubs", "needs_store": True, "attach_file": False,
     "reply": "This bills to: GL code: 180.3015.15917.000.6340.623020"},
    {"option": "Interior Building (Crow’s Nest)", "needs_store": False, "attach_file": False,
     "reply": "This bills to: GL code: 180.3015.10001.000.6340.623030.000.0000"},
    {"option": "NSO", "needs_store": False, "attach_file": False,
     "reply": "This is a NSO. This bills to: GL code: 170.3010.10125.000.6340.623050.000.0000"},
]

# --- Load Store List from Disk ---
if "store_df" not in st.session_state:
    if os.path.exists(STORE_LIST_PATH):
        st.session_state.store_df = pd.read_excel(STORE_LIST_PATH)
    else:
        st.session_state.store_df = None

# --- Store List Upload ---
st.header("📂 Upload or Replace Store List")
store_file = st.file_uploader("Upload your store list (Excel format)", type=["xlsx"], key="store_list_upload")
if store_file:
    try:
        store_df = pd.read_excel(store_file)
        store_df.to_excel(STORE_LIST_PATH, index=False)
        st.session_state.store_df = store_df
        st.success("✅ Store list uploaded and saved successfully.")
        st.dataframe(store_df.head())
    except Exception as e:
        st.error(f"Error reading file: {e}")

# --- Require Store List ---
if st.session_state.store_df is None:
    st.warning("Please upload a store list to continue.")
    st.stop()

# --- Single PDF Upload ---
st.header("📑 Upload Invoice and Assign Response")
pdf_file = st.file_uploader("Upload invoice PDF", type=["pdf"])

if pdf_file:
    with st.expander(f"📄 {pdf_file.name}", expanded=True):
        try:
            images = convert_from_bytes(pdf_file.read(), first_page=1, last_page=1)
            if images:
                st.image(images[0], caption="Page 1 Preview", use_container_width=True)
        except Exception as e:
            st.warning(f"Could not generate image preview: {e}")

        option_labels = [tpl["option"] for tpl in RESPONSE_TEMPLATES]
        selected_option = st.selectbox("Select response type", option_labels)
        tpl = next((tpl for tpl in RESPONSE_TEMPLATES if tpl["option"] == selected_option), None)

        store_input = None
        if tpl and tpl["needs_store"]:
            store_input = st.text_input("Enter store number(s)")

        if st.button("Generate Email Response"):
            from email.message import EmailMessage
            from pathlib import Path
            reply = tpl["reply"]
            attach_file = tpl.get("attach_file", False)
            df = st.session_state.store_df

            # Normalize column names for safety
            df.columns = [col.lower().strip() for col in df.columns]

            if tpl["needs_store"]:
                numbers = [s.strip() for s in store_input.replace(",", " ").split() if s.strip()]
                if len(numbers) == 1:
                    if "store number" in df.columns:
                        row = df[df["store number"].astype(str) == numbers[0]]
                        if not row.empty:
                            region = str(row.iloc[0]["region code"])
                            reply = reply.replace("XXXXX", region)
                        else:
                            reply = f"No matching store found for store number: {numbers[0]}"
                    else:
                        reply = "The store list does not contain a 'store number' column."
                elif len(numbers) > 1:
                    if "store number" in df.columns:
                        filtered = df[df["store number"].astype(str).isin(numbers)]
                        if not filtered.empty and attach_file:
                            output = io.BytesIO()
                            filtered.to_excel(output, index=False)
                            output.seek(0)
                            st.dataframe(filtered)
                            st.download_button("Download Store List", output.getvalue(), file_name=f"filtered_{pdf_file.name.replace('.pdf', '.xlsx')}")
                        reply = reply.replace("XXXXX", "multiple stores")
                    else:
                        reply = "The store list does not contain a 'store number' column."

            elif attach_file:
                output = io.BytesIO()
                df.to_excel(output, index=False)
                st.download_button("Download Full Store List", output.getvalue(), file_name=f"allstores_{pdf_file.name.replace('.pdf', '.xlsx')}")

            st.text_area("Email Body", value=reply, height=150)

            # Always prepare EML download
            email_msg = EmailMessage()
            email_msg["Subject"] = f"Response to: {pdf_file.name}"
            email_msg["To"] = "DSao@techstyle.com"
            email_msg.set_content(reply)

            if tpl["attach_file"]:
                if tpl["needs_store"] and len(numbers) > 1 and not filtered.empty:
                    with io.BytesIO() as output:
                        filtered.to_excel(output, index=False)
                        output.seek(0)
                        email_msg.add_attachment(output.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=f"filtered_{pdf_file.name.replace('.pdf', '.xlsx')}")
                elif not tpl["needs_store"]:
                    with io.BytesIO() as output:
                        df.to_excel(output, index=False)
                        output.seek(0)
                        email_msg.add_attachment(output.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=f"allstores_{pdf_file.name.replace('.pdf', '.xlsx')}")

            email_msg.add_attachment(pdf_file.getvalue(), maintype="application", subtype="pdf", filename=pdf_file.name)
            eml_output = io.BytesIO(email_msg.as_bytes())
            eml_filename = pdf_file.name.replace(".pdf", "_reply.eml")
            st.download_button("📩 Download .eml Draft", eml_output.getvalue(), file_name=eml_filename)

            if st.button("📤 Create Outlook Draft"):
                import subprocess
                import tempfile
                from pathlib import Path

                # Save the attachment temporarily if it exists
                attachment_path = ""
                if tpl["attach_file"]:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                        if tpl["needs_store"] and len(numbers) > 1 and not filtered.empty:
                            filtered.to_excel(tmp_file.name, index=False)
                        elif not tpl["needs_store"]:
                            df.to_excel(tmp_file.name, index=False)
                        attachment_path = tmp_file.name

                # Save the original PDF
                pdf_path = ""
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pdf_temp:
                    pdf_temp.write(pdf_file.getvalue())
                    pdf_path = pdf_temp.name

                import json
                safe_reply = json.dumps(reply)[1:-1]
                safe_subject = f"Response to: {pdf_file.name}"

                # AppleScript (for local use only)
                script = f'''
                tell application "Microsoft Outlook"
                    set newMessage to make new outgoing message with properties {{subject:"{safe_subject}", content:"{safe_reply}"}}
                    make new recipient at newMessage with properties {{email address:{{name:"D S", address:"DSao@techstyle.com"}}}}
                    tell newMessage
                        if "{attachment_path}" is not "" then
                            make new attachment with properties {{file:"{attachment_path}"}}
                        end if
                        if "{pdf_path}" is not "" then
                            make new attachment with properties {{file:"{pdf_path}"}}
                        end if
                    end tell
                    open newMessage
                end tell
                '''

                try:
                    subprocess.run(["osascript", "-e", script])
                except Exception as e:
                    st.warning(f"AppleScript execution skipped or failed: {e}")

                # EML export (for web use)
                from email.message import EmailMessage
                email_msg = EmailMessage()
                email_msg["Subject"] = safe_subject
                email_msg["To"] = "DSao@techstyle.com"
                email_msg.set_content(reply)

                # Attach files if they exist
                if attachment_path and Path(attachment_path).exists():
                    with open(attachment_path, "rb") as f:
                        email_msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=Path(attachment_path).name)

                if pdf_path and Path(pdf_path).exists():
                    with open(pdf_path, "rb") as f:
                        email_msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=Path(pdf_path).name)

                eml_output = io.BytesIO(email_msg.as_bytes())
                eml_filename = pdf_file.name.replace(".pdf", "_reply.eml")
                st.download_button("📩 Download .eml Draft", eml_output.getvalue(), file_name=eml_filename)
