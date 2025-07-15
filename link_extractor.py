import streamlit as st
import pandas as pd
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup

def load_eml_file(file_obj):
    msg = BytesParser(policy=policy.default).parse(file_obj)

    sender = msg['from']
    subject = msg['subject']
    date = msg['date']
    html_content = None

    for part in msg.walk():
        if part.get_content_type() == 'text/html':
            html_content = part.get_content()
            break

    return {
        "sender": sender,
        "subject": subject,
        "date": date,
        "html": html_content or ""
    }

def extract_links_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    links = []
    for a_tag in soup.find_all('a', href=True):
        link_text = a_tag.get_text(strip=True)
        href = a_tag['href']
        links.append({"link_text": link_text, "url": href})
    return links

def main():
    st.title("Outlook .eml Link Extractor")
    uploaded_file = st.file_uploader("Upload an .eml file", type=["eml"])

    if uploaded_file is not None:
        try:
            email_data = load_eml_file(uploaded_file)
            links = extract_links_from_html(email_data["html"])
            df = pd.DataFrame(links)

            st.subheader("Extracted Links")
            st.dataframe(df)

            if not df.empty:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button("Download CSV", csv, "extracted_links.csv", "text/csv")
        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
