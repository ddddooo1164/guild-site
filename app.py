ValueError: This app has encountered an error. The original error message is redacted to prevent data leaks. Full error details have been recorded in the logs (if you're on Streamlit Cloud, click on 'Manage app' in the lower right of your app).
Traceback:
File "/mount/src/guild-site/app.py", line 466, in <module>
    st.session_state.db_data = convert_sheets_to_dict(m_df, f_df)
                               ~~~~~~~~~~~~~~~~~~~~~~^^^^^^^^^^^^
File "/mount/src/guild-site/app.py", line 424, in convert_sheets_to_dict
    "gold": int(pd.to_numeric(row.get('gold', 0), errors='coerce') or 0),
            ~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
