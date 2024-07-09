angular.module('umbraco').controller('customExcelReportUploadController', ['$scope', '$http', function ($scope, $http) {
    $scope.allowedFileTypes = ['.xls', '.xlsx'];
    $scope.fileUploaded = false;

    $scope.uploadFile = function () {
        var file = document.getElementById("fileInput").files[0];

        if (!file) {
            alert("Please select a file to upload.");
            return;
        }

        var url = '/umbraco/surface/CustomExcelReportUploadApi/Upload';
        var xhr = new XMLHttpRequest();
        var formData = new FormData();

        formData.append("file", file);

        xhr.open('POST', url, true);

        xhr.onload = function () {
            debugger;
            if (xhr.status === 200) {
                var response = JSON.parse(xhr.responseText);
                if (response.success) {
                    console.log("File uploaded successfully:", response);
                    $scope.$apply(function () {
                     /*   $scope.fileUrl = response.fileUrl;*/
                        $scope.fileUploaded = true;
                    });
                } else {
                    console.error("Error uploading file:", response.message);
                    alert("Error uploading file: " + response.message);
                }
            } else {
                console.error("Error uploading file:", xhr.statusText);
                alert("Error uploading file. Please try again.");
            }
        };

        xhr.onerror = function () {
            console.error("Error uploading file:", xhr.statusText);
            alert("Error uploading file. Please try again.");
        };

        xhr.send(formData);
    };
}]);
