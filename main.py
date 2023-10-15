import streamlit as st  # at terminal, type: streamlit run appname.py      // pip install streamlit
from streamlit_option_menu import option_menu  # pip install streamlit_option_menu
import QAP
import SAWT
import SLSP
# import form_1604C
# import form_1604E
# import form_1604F
import form_2307


hide_menu = """
                <style> # MainMenu { visibility:hidden; }
                            footer { visibility:hidden; }
                            span {visibility: hidden; }
                </style>
            """

# basta naa ni nga setup (set_page_config) - streamlit display pages folder file
st.set_page_config(
    page_title="Alpha List",
    page_icon="book",

    # https://icons.getbootstrap.com/           # icons are coming from this link
)


def main():
    st.markdown(hide_menu, unsafe_allow_html=True)

    with st.sidebar:
        app = option_menu(
            menu_title='DLT ALPHA LIST',
            menu_icon='journal-text',
            options=['form_2307', 'SAWT', 'SLSP', 'QAP'],
            icons=['book', 'journal', 'journal-richtext', 'journal-text',
                   'journals', 'journal-check', 'card-list', 'card-checklist'],
            default_index=0,
            styles={
                "container": {"padding": "5!important", "background-color": 'grape'},
                "icon": {"color": "white", "font-size": "17px"},
                "nav-link": {"color": "white", "font-size": "12px", "text-align": "left", "margin": "0px"},
                "nav-link-selected": {"background-color": "teal"},
            }
        )

    if app == 'form_2307':
        form_2307.app()
    if app == 'QAP':
        QAP.app()
    if app == 'SAWT':
        SAWT.sawt_user_input_path()
    if app == 'SLSP':
        SLSP.app()


if __name__ == '__main__':
    main()
