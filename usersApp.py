import pandas as pd
import streamlit as st
from st_login_form import login_form
from st_supabase_connection import SupabaseConnection

# Initialize connection.
conn = st.connection("supabase",type=SupabaseConnection)

# Perform query.
rows = conn.query("*", table="mytable", ttl="10m").execute()

# Print results.
for row in rows.data:
    st.write(f"{row['name']} has a :{row['pet']}:")

st.title("Title")

client = login_form()

if st.session_state["authenticated"]:
    if st.session_state["username"]:
        st.success(f"Welcome {st.session_state['username']}")
    else:
        st.success("Welcome guest")
else:
    st.error("Not authenticated")