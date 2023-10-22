# VBA_Private-GPT (Work in Progress)
 Access company hosted private ChatGPT instance with SSO

Uses VBA-Web Tools (https://github.com/VBA-tools) for http:// requests in VBA to access your company's private instance of ChatGPT using microsoft SSO. 
My company uses Microsoft and Walk.Me as authentication services. Modify the code to fit your needs. 



SSO_GPT.bas Explanation
GitHub Copilot: This is a Visual Basic for Applications (VBA) module that provides a function called `GetConversationResponse`. This function takes an array of messages as input and returns a response from a chatbot API. The chatbot API is accessed using the `WebClient` class from the VBA-Web library. The `GetConversationResponse` function first calls the `GetTokenAndUserId` function to authenticate with the chatbot API. If authentication fails, the function returns an error message. If authentication succeeds, the function builds a request URL and payload, sends the request to the chatbot API, and parses the response JSON to extract the message.

The `GetTokenAndUserId` function checks if the access token and user ID have already been obtained. If not, the function calls the `GetSsoToken` function to obtain a single sign-on (SSO) token from the company's authentication server. The function then sends a request to the chatbot API to obtain the access token and user ID using the SSO token. If the request fails, the function returns `False`. If the request succeeds, the function stores the access token and user ID and returns `True`.

The module also defines several helper functions. The `GetSsoToken` function is a placeholder that assumes the existence of a function with that name that returns the SSO token as a string. The `AuthorizationHeader` function takes an access token as input and returns a string that can be used as an HTTP authorization header. The `UserIdHeader` function takes a user ID as input and returns a string that can be used as a custom HTTP header.

To use this module, you need to add a reference to the "Microsoft Scripting Runtime" and "VBA-Web" libraries under Tools > References in the VBA editor. You also need to replace the placeholders in the code with your own values. Specifically, you need to replace `<your-company>` with the name of your company, replace `<your-sso-token>` with your SSO token, and replace the `GetSsoToken` function with your own implementation. 

To improve the readability of this code, you could add comments to explain the purpose of each function and variable. You could also extract the URL of the chatbot API into a constant variable to make it easier to change in the future. To improve the performance of this code, you could consider caching the access token and user ID so that they don't need to be obtained on every call to `GetConversationResponse`.


