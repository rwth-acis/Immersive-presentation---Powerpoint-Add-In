
(function () {
    "use strict";

    var messageBanner;

    var user;

    var token;

    var tokenExp;

    var presentationList;

    var selectedPresentationId;

    // Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Benachrichtigungsmechanismus initialisieren und ausblenden
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            $('.loggedIn').hide();

            $('#button-login').click(login);
            $('#logout-button').click(logout);
            $('#presentation-select').on('change', presentationSelectChanged);

            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    function login() {

        let email = $('#login-email').val();
        let pwd = $('#login-pwd').val();
        let loginSettings = {
            "url": "https://cloud19.dbis.rwth-aachen.de/auth/login",
            "method": "POST",
            "timeout": 0,
            "headers": {
                "Content-Type": "application/x-www-form-urlencoded"
            },
            "data": {
                "email": email,
                "password": pwd
            },
            success: function (returnData) {
                //extract login respone information
                user = returnData.user;
                token = returnData.token;
                tokenExp = returnData.exp;

                //prepare logged in view
                $('#username').text(user.email);

                //switch view to logged in
                $('.loggedOut').hide();
                $('.loggedIn').show();

                showNotification('Login Successfull! ', 'Welcome ' + returnData.user.email);

                loadPresentationList();
            },
            error: function (xhr, status, error) {
                var errorMessage = xhr.status + ': ' + xhr.statusText
                console.log('Error - ' + errorMessage);
                showNotification('Login Error: ', errorMessage);
            }
        };

        $.ajax(loginSettings);
    }

    function logout() {
        $('.loggedIn').hide();
        $('.loggedOut').show();
    }

    // Liest Daten aus der aktuellen Dokumentauswahl und zeigt eine Benachrichtigung an.
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('Der ausgewählte Text lautet:', '"' + result.value + '"');
                } else {
                    showNotification('Fehler:', result.error.message);
                }
            }
        );
    }

    // Eine Hilfsfunktion zum Anzeigen von Benachrichtigungen.
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function loadPresentationList() {

        var presentationListSettings = {
            "url": "https://cloud19.dbis.rwth-aachen.de/presentations",
            "method": "GET",
            "timeout": 0,
            "headers": {
                "Authorization": "Bearer " + token
            }
            ,
            success: function (returnData) {
                //extract login respone information
                presentationList = returnData.presentations;

                //Add presentation names to the dropdown
                $.each(presentationList, function (i, presentation) {
                    $('#presentation-select').append($('<option>', {
                        value: presentation.idpresentation,
                        text: presentation.name
                    }));
                });
                //update currently selected because adding new options to a select does not fire a change event
                presentationSelectChanged();
            },
            error: function (xhr, status, error) {
                var errorMessage = xhr.status + ': ' + xhr.statusText
                console.log('Error - ' + errorMessage);
                showNotification('Error Loading owned Presentations: ', errorMessage);
            }
        };

        $.ajax(presentationListSettings).done(function (response) {
            console.log(response);
        });
    }

    function presentationSelectChanged() {
        selectedPresentationId = $('#presentation-select').val();
    }

    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }

})();
