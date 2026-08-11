# Testing_ERP_Systems
This repository provides a framework for developing profiles and resumes that can be used to test ERP systems like UkG or Workday. 

Data set generation:

Recreate these artifacts using the following Colab notebooks and your own organization's current pool of open positions. The included example dataset uses a fictional set of higher-education job postings as the starting point in the ApplicantData.ipynb file.

ApplicantData.ipynb

This notebook uses python to generate a random list of applicants to assist in generating candidate records for testing an ERP system prior to go live. This process could be improved with additional data instead of hard coded random lists. This data was created to generate word and pdf resumes to be uploaded alongside candidate records. View the [ResumeGenerator.ipynb](ResumeGenerator.ipynb) notebook once the test set is generated here to generate your own resumes.


ResumeGenerator.ipnyb

In this notebook, I provide a method for generating resume artifacts that can be used when conducting testing on human resources and enterprise resource planning (ERP) systems. I construct the underlying data used in this resume generator using a different notebook so please review the [ApplicantData.ipynb](ApplicantData.ipynb) if you would like to generate your own test corpus. 



