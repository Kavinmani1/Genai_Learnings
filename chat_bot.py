## we can also do the deployment in huggingface spaces

import streamlit as st
from langchain.llms import OpenAI

from langchain.schema import HumanMessage,SystemMessage,AIMessage
from langchain.chat_models import ChatOpenAI

from dotenv import load_dotenv #load_dotenv reads key-value pairs from a .env file and can set them as environment variables.
import os

## Streamlit UI
st.set_page_config(page_title="Conversational Q&A Chatbot")
st.header("Hey, Let's Chat")


os.environ["OPENAI_API_KEY"] = "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"
llm=OpenAI(openai_api_key=os.environ["OPENAI_API_KEY"],temperature=0.6)

chat=ChatOpenAI(temperature=0.5)


if 'flowmessages' not in st.session_state:
    st.session_state['flowmessages']=[
        SystemMessage(content="Yor are a comedian AI assitant")
    ]

## Function to load OpenAI model and get respones

def get_chatmodel_response(question):

    st.session_state['flowmessages'].append(HumanMessage(content=question))
    answer=chat(st.session_state['flowmessages'])
    st.session_state['flowmessages'].append(AIMessage(content=answer.content))
    return answer.content

input=st.text_input("Input: ",key="input")
response=get_chatmodel_response(input)

submit=st.button("Ask the question")

## If ask button is clicked

if submit:
    st.subheader("The Response is")
    st.write(response)
