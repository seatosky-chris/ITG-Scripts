# Constants

# Find the nuget package location with "dotnet nuget locals global-packages -l" in a terminal
# Then naviggate to the "libphonenumber-csharp" folder and find the latest PhoneNumbers.dll for a version of .Net that will work on this system
# Use the full path for this constant
$phoneNumbersDLLPath = "C:\Users\Administrator\.nuget\packages\libphonenumber-csharp\8.12.34\lib\net46\PhoneNumbers.dll"

Add-Type -Path $phoneNumbersDLLPath

# The C# code for formatting phone numbers
Add-Type -TypeDefinition @"
using System;
using PhoneNumbers;
using System.Text.RegularExpressions;

public class PhoneNumberLookup 
{
    public static string FormatPhoneNumber(string phoneString)
    {
        PhoneNumberUtil phoneNumberUtil = PhoneNumbers.PhoneNumberUtil.GetInstance();

        string parsedPhoneString;
        if (phoneString.Trim().StartsWith("+") || phoneString.Trim().StartsWith("011")) 
        {
            parsedPhoneString = phoneString.Trim();
        } 
        else 
        {
            string extension = "";
        
            if (Regex.IsMatch(phoneString, "[a-zA-Z]")) 
            {
                Match letterMatches = Regex.Match(phoneString, "[a-zA-Z]");
                extension = " " + phoneString.Substring(letterMatches.Index);
                parsedPhoneString = phoneString.Substring(0, letterMatches.Index);
                parsedPhoneString = Regex.Replace(parsedPhoneString, "\\D", "");
            } 
            else 
            {
                parsedPhoneString = Regex.Replace(phoneString, "\\D", "");
            }
            
            if (parsedPhoneString.Length == 10) 
            {
                parsedPhoneString = "+1 " + parsedPhoneString;
            } 
            else if (parsedPhoneString.Length > 10) 
            {
                parsedPhoneString = "+" + parsedPhoneString;
            }

            parsedPhoneString += extension;
        }

        PhoneNumber phoneNumber = phoneNumberUtil.Parse(parsedPhoneString, "CA");

        if (phoneNumber.CountryCode == 1) 
        {
            return phoneNumberUtil.Format(phoneNumber, PhoneNumbers.PhoneNumberFormat.NATIONAL);
        } 
        else 
        {
            //return phoneNumberUtil.Format(phoneNumber, PhoneNumbers.PhoneNumberFormat.INTERNATIONAL);
            return phoneNumberUtil.FormatOutOfCountryCallingNumber(phoneNumber, "CA");
        }
    }
}
"@ -ReferencedAssemblies @($phoneNumbersDLLPath, "System.Text.RegularExpressions")

# Example
$phoneNumberTests = @("2024561111", "83835", "604.291.1255 ext. 82825", "011-65-9619-6676", "+16047880877", "+1 2502121767", "+44 (0) 118 951 9565", "+1 (236) 427-5574", "604.925.6665 ex236", "604.540.0029 Ext 110", "604.455.0366x132", "604.291.1255 ext. 82825", "250-861-1515 ext.218");

foreach ($phoneString in $phoneNumberTests) {

    if ($phoneString.StartsWith("011 ") -or $phoneString -match "^\(\d\d\d\) " -or $phoneString.Trim() -eq "email") {
        $formattedPhoneNumber = $phoneString.Trim();
        Write-Host "# already formatted.";
    } else {
        try {
            $formattedPhoneNumber = [PhoneNumberLookup]::FormatPhoneNumber($phoneString);
        } catch {
            $formattedPhoneNumber = $phoneString;
        }
    }

    Write-Host "Formatted: $($formattedPhoneNumber)";
}