# Office 365 SharePoint file management

This sample demonstrates how to upload files to SharePoint Online using Office 365 Authentication and AAD bearer tokens in a multi-tenant application.


## How to Run this Sample
To run this sample, you need:

1. Visual Studio 2015
2. [Office Developer Tools For Visual Studio 2015 Update 1](http://aka.ms/officedevtoolsforvs2015)
3. [Office 365 Developer Subscription](http://dev.office.com/devprogram)

## Step 1: Clone or download this repository
From your Git Shell or command line:

`git clone https://github.com/dougperkes/O365-SharePoint-FilePush.git`

## Step 2: Build the Project
1. Open the project in Visual Studio 2015.
2. Simply Build the project to restore NuGet packages.
3. Ignore any build errors for now as we will configure the project in the next steps.

## Step 3: Configure the sample
Once downloaded, open the sample in Visual Studio.

### Register Azure AD application to consume Office 365 APIs
Office 365 applications use Azure Active Directory (Azure AD) to authenticate and authorize users and applications respectively. All users, application registrations, permissions are stored in Azure AD.

Using the Office 365 API Tool for Visual Studio you can configure your web application to consume Office 365 APIs. 

1. In the Solution Explorer window, **right click your project -> Add -> Connected Service**.
2. The Add Connected Service wizard dialog box will appear. Choose **Microsoft -> Office 365 APIs** and click **Configure**.
3. On the **Select Domain** page, enter your Azure AD developer tenant name, i.e. contosodev.onmicrosoft.com.
4. Click **Next** and you will be promted to enter your O365 developer subscription credentials if they have not already been cached by Visual Studio.
5. On the **Configure Application** page, select "Create a new Azure AD application to access Office 365 API services". Click **Next**
6. Click the **Sites** page and select **Run search queries** and **Read and write items in all site collections**
7. Click the **Users and Groups** page and select **Sign you in and read your profile**
8. Click Finish

After clicking Finish in the Add Connected Service dialog box, Office 365 client libraries (in the form of NuGet packages) for connecting to Office 365 APIs will be added to your project. 

In this process, Office 365 API tool registered an Azure AD Application in the Office 365 tenant that you signed in the wizard and added the Azure AD application details to web.config. 

### Step 4: Build and Debug your web application
Now you are ready for a test run. Hit F5 to test the app.

