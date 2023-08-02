package gst.test;
import java.util.*;

public class UrlLink {
	public static void main(String[] args)

	{
String clientId = "77543861055-r619n5omutegukcrb3irnghc0cea7mld.apps.googleusercontent.com";
String redirectUri = "http://localhost"; // The redirect URI should match the one set up in your OAuth client ID configuration

String authorizationUrl = "https://accounts.google.com/o/oauth2/auth?" +
        "client_id=" + clientId +
        "&redirect_uri=" + redirectUri +
        "&response_type=code" +
        "&scope=https://www.googleapis.com/auth/drive"; // Replace the scope with the required API scope(s) for your application

System.out.println("Authorization URL: " + authorizationUrl);

}
}