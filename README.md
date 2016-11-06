# Audit Analytics

In this repository are example audit analytic tests using the python ecosystem. 

## Python dependencies: 

* Python 2.7
* pyodbc
* pywin32 (for outlook email)
* pandas

I recommend installing Anaconda https://www.continuum.io/downloads as it provides a complete environment with the dependencies already installed and does not require a lot of manual installation.

## R dependencies:

The R program as written installs the three required dependencies: 
* RODBC
* xlsx
* sendmailR (for gmail email) 

through the use of a library named pacman. So, the only manual installing you need to do is when you turn on R for the first time, run the following code:

install.packages("pacman")

The above code installs pacman, which installs and loads the other libraries when you run the program. 
To download R, visit the following link:
https://cran.r-project.org

For any questions or troubleshooting, contact brassatc@me.com


