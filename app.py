"""Standalone Streamlit app for outsource attendance."""

from __future__ import annotations

import streamlit as st

from src.outsource_attendance import (
    render_attendance_admin_page,
    render_attendance_observer_page,
    render_outsource_login_page,
)


PAGES = {
    "Outsource Login": render_outsource_login_page,
    "Observer Desk": render_attendance_observer_page,
    "Admin Panel": render_attendance_admin_page,
}


def apply_white_theme() -> None:
    st.markdown(
        """
        <style>
        :root {
            color-scheme: light;
        }

        .stApp {
            background: #ffffff;
            color: #111827;
        }

        [data-testid="stSidebar"] {
            background: #f8fafc;
            border-right: 1px solid #e5e7eb;
        }

        [data-testid="stSidebar"] * {
            color: #111827;
        }

        h1, h2, h3, h4, h5, h6,
        .stMarkdown, .stCaption, label, p {
            color: #111827;
        }

        div[data-testid="stMetric"],
        div[data-testid="stForm"],
        div[data-testid="stExpander"] {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
        }

        div[data-testid="stForm"] {
            padding: 1rem;
        }

        .stDataFrame {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
        }

        div[data-baseweb="input"] > div,
        div[data-baseweb="select"] > div,
        textarea {
            background: #ffffff;
            border-color: #cbd5e1;
            color: #111827;
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 8px;
            border: 1px solid #2563eb;
            background: #2563eb;
            color: #ffffff;
            font-weight: 600;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: #1d4ed8;
            background: #1d4ed8;
            color: #ffffff;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="Outsource Attendance",
        page_icon="ATT",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_white_theme()

    with st.sidebar:
        st.title("Attendance")
        selected_page = st.radio(
            "Open page",
            options=list(PAGES.keys()),
            label_visibility="collapsed",
        )
        st.divider()
        st.caption("Outsource attendance system")

    PAGES[selected_page]()


if __name__ == "__main__":
    main()
