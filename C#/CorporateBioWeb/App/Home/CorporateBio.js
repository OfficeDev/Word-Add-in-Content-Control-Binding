/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
/// <reference path="../App.js" />


    // This function is run when the app is ready to start interacting with the host application
    // It ensures the DOM is ready before binding to named content controls, and adding click handlers to buttons
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Create a binding to the RichTextContentControl that has a title of EmployeeName
            Office.context.document.bindings.addFromNamedItemAsync("EmployeeName", Office.BindingType.Text, { id: 'nameBinding' }, function (asyncResult) {
                // Get a reference to the DIV in the CorporateBio.html page, because we will write some text to the DIV
                var report = document.getElementById("validationReport");
                var reportText;
                // Provide status information about the binding. NOTE: In a real solution, you wouldn't bother showing this to the user
                // but it can help you troubleshoot later if the binding cannot be found.
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    reportText = document.createTextNode("Could not bind to Employee Name. Error Details are: " + asyncResult.error.message);
                }
                else {
                    reportText = document.createTextNode("Binding to EmployeeName: Success");
                }
                report.appendChild(reportText);
            });

            // Create a binding to the RichTextContentControl that has a title of EmployeePosition
            Office.context.document.bindings.addFromNamedItemAsync("EmployeePosition", Office.BindingType.Text, { id: 'positionBinding' }, function (asyncResult) {
                // Get a reference to the DIV in the CorporateBio.html page, because we will write some text to the DIV
                var report = document.getElementById("validationReport");
                var reportText;
                // Provide status information about the binding. NOTE: In a real solution, you wouldn't bother showing this to the user
                // but it can help you troubleshoot later if the binding cannot be found.
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    reportText = document.createTextNode("Could not bind to Employee Position. Error Details are: " + asyncResult.error.message);
                }
                else {
                    reportText = document.createTextNode("Binding to Employee Position: Success");
                }
                // The following line break is added because this text should be on a new line (after the EmployeeName status)
                var lineBreak = document.createElement("br");
                report.appendChild(lineBreak);
                report.appendChild(reportText);
            });

            // Create a binding to the RichTextContentControl that has a title of EmployeeAboutMe
            Office.context.document.bindings.addFromNamedItemAsync("EmployeeAboutMe", Office.BindingType.Text, { id: 'aboutMeBinding' }, function (asyncResult) {
                // Get a reference to the DIV in the CorporateBio.html page, because we will write some text to the DIV
                var report = document.getElementById("validationReport");
                var reportText;
                // Provide status information about the binding. NOTE: In a real solution, you wouldn't bother showing this to the user
                // but it can help you troubleshoot later if the binding cannot be found.
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    reportText = document.createTextNode("Could not bind to About Me. Error Details are: " + asyncResult.error.message);
                }
                else {
                    reportText = document.createTextNode("Binding to About Me: Success");
                }
                // The following line break is added because this text should be on a new line (after the EmployeePosition status)
                var lineBreak = document.createElement("br");
                report.appendChild(lineBreak);
                report.appendChild(reportText);
            });

            // Wire up the click events of the two buttons in the CorporateBio.html page.
            $('#validateData').click(function () { validate(); });
            $('#submitData').click(function () { submit(); });

        });
    };
    // Class level variable that gets set by the validate() function, and read by the submit function. Both of the functions are below.
    var isValidName = false;
    var isValidPosition = false;
    var isValidAbout = false;

    // This function runs when the user clicks the [Validate data] button
    function validate() {
        // Get a reference to the DIV in the CorporateBio.html page, because we will write some text to the DIV
        var report = document.getElementById("validationReport");
        var reportText;
        //Remove all nodes from the validationReport <DIV> so we have a clean space to write to
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        // Attempt to get the nameBinding that should have been added in the $(document).ready function
        Office.context.document.bindings.getByIdAsync('nameBinding', function (asyncResult) {
            // If retrieving the binding fails, tell the user
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                reportText = document.createTextNode("Error retrieving the Employee Name binding");
                report.appendChild(reportText);
                //Set the flag to false, so that it can be read by the submit function
                this.isValidName = false;
            }
            else {
                //If retrieving the binding succeeds, attempt to get the actual data 
                Office.select("bindings#nameBinding").getDataAsync(function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        //If retrieving the data fails, tell the user
                        reportText = document.createTextNode("Error retrieving the Employee Name binding: " + asyncResult.error.message);
                        report.appendChild(reportText);
                        //Set the flag to false, so that it can be read by the submit function
                        this.isValidName = false;
                    } else {
                        // If retrieving the data succeeds, check whether the content control still contains its placeholder text.
                        if (asyncResult.value == "Click here to enter text.") {
                            //Data has not been entered, so tell the user
                            reportText = document.createTextNode("Employee Name is required!");
                            report.appendChild(reportText);
                            //Set the flag to false, so that it can be read by the submit function
                            this.isValidName = false;
                        }
                        else {
                            //Data has been entered, so provide a summary. 
                            //NOTE: In a real solution you probabaly wouldn't provide this summary,
                            //but it can help you see that the correct data is returned when running this sample
                            reportText = document.createTextNode("Employee Name: " + asyncResult.value);
                            report.appendChild(reportText);
                            //Set the flag to true, so that it can be read by the submit function
                            this.isValidName = true;
                        }
                    }
                });

            }
        });

        // Attempt to get the positionBinding that should have been added in the $(document).ready function
        Office.context.document.bindings.getByIdAsync('positionBinding', function (asyncResult) {
            var lineBreak = document.createElement("br");
            report.appendChild(lineBreak);
            // If retrieving the binding fails, tell the user
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                reportText = document.createTextNode("Error retrieving the Employee Position binding");
                report.appendChild(reportText);
                //Set the flag to false, so that it can be read by the submit function
                this.isValidPosition = false;
            }
            else {
                //If retrieving the binding succeeds, attempt to get the actual data 
                Office.select("bindings#positionBinding").getDataAsync(function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        //If retrieving the data fails, tell the user
                        reportText = document.createTextNode("Error retrieving the Employee Position binding: " + asyncResult.error.message);
                        report.appendChild(reportText);
                        //Set the flag to false, so that it can be read by the submit function
                        this.isValidPosition = false;
                    } else {
                        // If retrieving the data succeeds, check whether the content control still contains its placeholder text.
                        if (asyncResult.value == "Click here to enter text.") {
                            //Data has not been entered, so tell the user
                            reportText = document.createTextNode("Employee Position is required!");
                            report.appendChild(reportText);
                            //Set the flag to false, so that it can be read by the submit function
                            this.isValidPosition = false;
                        }
                        else {
                            //Data has been entered, so provide a summary. 
                            //NOTE: In a real solution you probabaly wouldn't provide this summary,
                            //but it can help you see that the correct data is returned when running this sample
                            reportText = document.createTextNode("Employee Position: " + asyncResult.value);
                            report.appendChild(reportText);
                            //Set the flag to true, so that it can be read by the submit function
                            this.isValidPosition = true;
                        }
                    }
                });

            }
        });

        // Attempt to get the aboutMeBinding that should have been added in the $(document).ready function
        Office.context.document.bindings.getByIdAsync('aboutMeBinding', function (asyncResult) {
            var lineBreak = document.createElement("br");
            report.appendChild(lineBreak);
            // If retrieving the binding fails, tell the user
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                reportText = document.createTextNode("Error retrieving the ABout Me binding");
                report.appendChild(reportText);
                //Set the flag to false, so that it can be read by the submit function
                this.isValidAbout = false;
            }
            else {
                //If retrieving the binding succeeds, attempt to get the actual data 
                Office.select("bindings#aboutMeBinding").getDataAsync(function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        //If retrieving the data fails, tell the user
                        reportText = document.createTextNode("Error retrieving the About Me binding: " + asyncResult.error.message);
                        report.appendChild(reportText);
                        //Set the flag to false, so that it can be read by the submit function
                        this.isValidAbout = false;
                    } else {
                        // If retrieving the data succeeds, check whether the content control still contains its placeholder text.
                        if (asyncResult.value == "Click here to enter text.") {
                            //Data has not been entered, so tell the user
                            reportText = document.createTextNode("About Me is required!");
                            report.appendChild(reportText);
                            //Set the flag to false, so that it can be read by the submit function
                            this.isValidAbout = false;
                        }
                        else {
                            //Data has been entered, so provide a summary. 
                            //NOTE: In a real solution you probabaly wouldn't provide this summary,
                            //but it can help you see that the correct data is returned when running this sample
                            reportText = document.createTextNode("About Me: " + asyncResult.value);
                            report.appendChild(reportText);
                            //Set the flag to false, so that it can be read by the submit function
                            this.isValidAbout = true;
                        }
                    }
                });
            }
        });

    }

    // This function runs when the user clicks the [Submit data] button
    function submit() {
        // Remove all nodes from the validationReport <DIV> so we have a clean space to write to
        var report = document.getElementById("validationReport");
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }
        // this.isValidName and this.isValidPosition and this.isValidAbout all have a default of false,
        // so if the [Validate data] button 
        // has not yet been clicked we'll know not to submit data.
        // If the [Validate data] button HAS been clicked, then these values will only be true
        // if each binding has appropriate data. See the validate() function above for details
        if ((this.isValidName) && (this.isValidPosition) && (this.isValidAbout)) {
            // Add your own code here for saving the values to a SharePoint list, 
            // or SQL Server Database, or perhaps passing them through a SharePoint workflow
            // and then inform the user as follows:
            var reportText = document.createTextNode("Data has been submitted. Thank you!");
            report.appendChild(reportText);
        }
        else {
            // Inform the user that they must validate before submitting
            var reportText = document.createTextNode("Please validate before submitting");
            report.appendChild(reportText);
        }
    }

// *********************************************************
//
// Word-Add-in-Content-Control-Binding, https://github.com/OfficeDev/Word-Add-in-Content-Control-Binding
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
