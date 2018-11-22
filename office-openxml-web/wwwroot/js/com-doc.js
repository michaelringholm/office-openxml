$(function() {
    console.log( "ready!" );
    getTemplates();
    $("#btnCreatePDF").click(function() {createPDF();})    
});

function getTemplates() {
    console.log( "Getting templates!" );
    $.ajax({
        url: "/api/doc/templates", 
        success: function(blobItems) {
            for(var i=0;i<blobItems.length;i++)
                $('#templates').append($("<option></option>").attr("value", blobItems[i]).text(blobItems[i]));
        }
    });
}

function createPDF() {
    console.log( "Creating PDF!" );
    var blobItemName = $("#templates").val();
    var bodyText = $("#bodyText").val();
    var data = { blobItemName: blobItemName, bodyText: bodyText };

    $.ajax({
        url: "/api/doc/modify",
        type: "POST",
        contentType: "application/json",
        data: JSON.stringify(data),
        success: function(pdfGuid) {
            $("#pdfLink").html('<a id="pdfGuid" href="/api/doc/preview/' + pdfGuid + '">Preview</a>');
        }
    });    
}