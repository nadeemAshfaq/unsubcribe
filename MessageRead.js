var app = angular.module('MsSignIn', ['ngMaterial'], function ($mdThemingProvider) {
    $mdThemingProvider.theme('default')
        .primaryPalette('blue', {
            'default': '500', // by default use shade 400 from the pink palette for primary intentions

        });
});
app.controller('SignInCtrl', function ($scope) {


    var Logindialog;
    var token = "";
    Office.onReady(function () {

        ProgressCirActive();

      //  var GetStoreToken = window.localStorage.getItem("getlocal");
        //console.log(GetStoreToken);

    //    if (GetStoreToken) {
            $scope.Login_Btn = true;
            $scope.MainPage = false;
            $scope.Log_Out = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
     //   } else {
            $scope.Login_Btn = false;
            $scope.MainPage = true;
            $scope.Log_Out = true;
            if (!$scope.$$phase) {
                $scope.$apply();
            };
     //   };


        function LogprocessMessage(arg) {
            Logindialog.close();
            token = arg.message;
      

                $scope.Login_Btn = true;
                $scope.MainPage = false;
                $scope.Log_Out = false;
                if (!$scope.$$phase) {
                    $scope.$apply();
                };
        

        };



        $scope.email = Office.context.mailbox.item.from.emailAddress;




        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                //  console.log(asyncResult.value);
                var Contnt = asyncResult.value;
                var urlRegex = /(https?:\/\/[^\s]+)/g;

                if (Contnt.match(urlRegex)) {

                    Contnt.replace(urlRegex, function (url) {
                        if (url.indexOf('unsubscribe') > -1) {
                            $scope.URLofUnsubcribe = url;
                            ProgressCirInActive();
                        }

                        else {


                            ProgressCirInActive();

                        }
                    });

                }
                else {
                    ProgressCirInActive();
                }

             

                // console.log(arry)
            }
        });

       

        $scope.SignInMS = function () {

            ////////////////Correct////////////////
           // var link = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=60a473c1-1377-453c-8b5b-0d26a793ba6d&response_type=token&redirect_uri=https://nadeemashfaq.github.io/unsubcribe/Templates/redirectPage.html&scope=Mail.Send&response_mode=fragment&state=12345&nonce=678910";
           //var link = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=b916048b-f4dc-429d-9847-677d34bdbed7&response_type=token&redirect_uri=https://nadeemashfaq.github.io/unsubcribe/Templates/redirectPage.html&scope=Mail.Send&response_mode=fragment&state=12345&nonce=678910";
            var link = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=d4d7d3ce-79df-4d3f-b9da-95b65654fad8&response_type=token&redirect_uri=https://nadeemashfaq.github.io/unsubcribe/Templates/redirectPage.html&scope=Mail.Send&response_mode=fragment&state=12345&nonce=678910";
       //     var link = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=b916048b-f4dc-429d-9847-677d34bdbed7&response_type=token&redirect_uri=https://localhost:44382/Templates/redirectPage.html&scope=Mail.Send&response_mode=fragment&state=12345&nonce=678910";


            Office.context.ui.displayDialogAsync(link, { height: 50, width: 30 },
                function (asyncResult) {
                    Logindialog = asyncResult.value;
                    Logindialog.addEventHandler(Office.EventType.DialogMessageReceived, LogprocessMessage);
                });
        };


        function getItemRestId() {
            if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
                // itemId is already REST-formatted.
                return Office.context.mailbox.item.itemId;
            } else {
               
                return Office.context.mailbox.convertToRestId(
                    Office.context.mailbox.item.itemId,
                    Office.MailboxEnums.RestVersion.v2_0
                );
            }
        }


       // console.log(Office.context.mailbox.item);



        $scope.Reply = function () {

            var ITEM = Office.context.mailbox.item;
            console.log(ITEM);
            console.log(ITEM.itemId);
          //  console.log(GetStoreToken);
            
            var Item_id = getItemRestId();
          //  var Item_id = ITEM.conversationId;


            var Resply = {
                "message": {
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": ITEM.from.emailAddress,
                                "name": ITEM.from.displayName
                            }
                        },
                     
                    ]
                },
                "comment": "again Testing"
            }
                
            $.ajax({
                url: "https://graph.microsoft.com/v1.0/me/messages/" + Item_id + "/reply",
                method:"POST",
                headers: {
                    "Authorization": "Bearer " + token,
                    "Content-Type": "application/json"
                },
                data: JSON.stringify(Resply),
                success: function (result) {
                    console.log(result);
                },
                error: function (error) {
                    console.log(error);
                }

            });

        }





        $scope.RemoveLocal = function () {
          
                $scope.Login_Btn = false;
                $scope.MainPage = true;
            $scope.Log_Out = true;


                if (!$scope.$$phase) {
                    $scope.$apply();
                };
          
        };




     


        function ProgressCirActive() {
            $("#StartProgressCir").show(function () {

                $("#ProgressBgDivCir").show();
            
                $scope.ddeterminateValue = 15;
                $scope.showProgressLinear = false;
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            });
        };

        function ProgressCirInActive() {
            $("#StartProgressCir").hide(function () {
                setTimeout(function () {
                    $scope.ddeterminateValue = 0;
                    $scope.showProgressLinear = true;
                    $("#ProgressBgDivCir").hide();
                   
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }, 500);
            });
        };

    });
});