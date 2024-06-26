﻿using System;
using PhoneNumbers;
using System.Text.RegularExpressions;

public class Program
{
    public static void Main()
    {
        // Test
        string phoneString = "+16047880877";
        string formattedPhoneNumber;
        Match cadPrefix = Regex.Match(phoneString, @"^\(\d\d\d\) ");

        if (phoneString.StartsWith("011 ") || (cadPrefix.Success && cadPrefix.Index == 0) || phoneString.Trim() == "email") 
        {
            formattedPhoneNumber = phoneString.Trim();
        } 
        else 
        {
            try 
            {
                formattedPhoneNumber = PhoneNumberLookup.FormatPhoneNumber(phoneString);
            } 
            catch 
            {
                formattedPhoneNumber = phoneString;
            }
        }

        Console.WriteLine($"Formatted: {formattedPhoneNumber}");
    }
}

// Primary Class
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