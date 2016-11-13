Office.initialize = function (reason) {
	$(document).ready(function () {
		jQuery(function () {
			$("#dialog").dialog();
		});
		jQuery('#form1').submit(function (e) {
			e.preventDefault();
			var docData = jQuery('#form1 .field1').val();
			var coerc = jQuery('#form1 .field2').val();
			writeData(docData, coerc);
		});
		jQuery('#form2').submit(function (e) {
			e.preventDefault();
			var paneData = jQuery('#form2 .field1').val();
			var coerc = jQuery('#form1 .field2').val();
			printData(paneData, coerc);
		});
		jQuery('#form4').submit(function (e) {
			e.preventDefault();
			var theID = jQuery('#form4 .field1').val();
			createBinding(theID);
		});
		jQuery('#form5').submit(function (e) {
			e.preventDefault();
			var theID = jQuery('#form5 .field1').val();
			var newData = jQuery('#form5 .field2').val();
			var coerc = jQuery('#form5 .field3').val();
			updateBoundText(newData, theID, coerc);
		});
		jQuery('#form6').submit(function (e) {
			e.preventDefault();
			var theID = jQuery('#form6 .field1').val();
			track(theID);
		});
		jQuery('#form7').submit(function (e) {
			e.preventDefault();
			var theID = jQuery('#form7 .field1').val();
			changeTrack(theID);
		});
		jQuery('#form8').submit(function (e) {
			e.preventDefault();
			var alias = jQuery('#form8 .field3').val();
			var newID = jQuery('#form8 .field1').val();
			bindContentControl(alias, newID);
		});
		$('#greyify').click(function () {
			$('body').toggleClass('greyify');
		});
		jQuery('#postTest').click(function () {
			jQuery.ajax({
				type: 'GET',
				url: 'https://public.opencpu.org/ocpu/library/MASS/data/DDT/json',
				complete: function (data) {
					ajaxCallback(data);
				}
			});
		});
    WordToPaneXMLListener();
    
		$("#goTo").click(goToByID);
		$("#theJ").dialog({
			autoOpen: false
		});
		$("#jQueryModal").on("click", function () {
			$("#theJ").dialog("open");
		});
		var hbTempSrc1 = jQuery("#hbTempSrc1").html();
		var hbTemplate1 = Handlebars.compile(hbTempSrc1);
		var trackI = 1;
		var trackC = 1;
		var hbData = {
			tagOption: [
        "Advisories",
        "Binding",
        "Declarations",
        "Financial",
        "International",
        "Non-Binding",
        "Notices",
        "Requests",
        "Waivers",
        "Warnings"
      ]
		}
	});
}

var doge = [['Woof'], ['*Much Stare Amazed*', , 'Multiple Columns Supported'], ['*Very Still Stares*']];
var htmlVar = '<span style=\'color:red;\'>I am a span of red text formatted using html inline css with a span element.</span><br/><span style=\'background-color:blue;\'>And I am a highlighted area!</span>';
var asyncRibbonVar = '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon startFromScratch="false"><tabs><!-- Your tabs and controls go here --><tab id="CustomTab" label="asyncTab"><group id="SampleGroup" label="Sample Group"></group></tab></tabs></ribbon></customUI>';
var trackI = 1;
var trackC = 0;

function ajaxCallback(data) {
	jQuery("#results").text(data.responseText);
}

function writeData(data, coerc) {
	Office.context.document.setSelectedDataAsync(data, {
		coercionType: coerc
	});
}

function readData(coerc) {
	Office.context.document.getSelectedDataAsync(coerc, function (selec) {
		printData(selec.value);
	});
}

function printData(data, slice) {
	{
		var printOut = "";
		if (slice === 'yes') {
			for (var x = 0; x < data.length; x++) {
				for (var y = 0; y < data[x].length; y++) {
					printOut += data[x][y] + ",";
				}
			}
		} else {
			printOut = data;
		}
		document.getElementById("results").innerText = printOut;
	}
}

function createBinding(theID) {
	Office.context.document.bindings.addFromSelectionAsync("text", {
		id: theID
	}, function (asyncResult) {
		if (asyncResult.status == "failed") {
			printData("Action failed with error: " + asyncResult.error.message);
		} else {
			printData("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id);
		}
	});
}

function updateBoundText(txt, theID, coerc) {
	//Go to binding by id.
	Office.select("bindings#" + theID).setDataAsync(txt, {
		coercionType: coerc
	}, function (asyncResult) {});
}

function track(theID) {
	printData('Attempting to nav track object with id = ' + theID);
	Office.select("bindings#" + theID).addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelChange);
}

function changeTrack(theID) {
	printData('Attempting to change track object with id = ' + theID);
	Office.select("bindings#" + theID).addHandlerAsync(Office.EventType.BindingDataChanged, changeTrackWriter);
}

function changeTrackWriter(args) {
  printData(args.binding.id + " has been modified. Mod number: " + trackC);
	trackC++;
}

function onBindingSelChange(args) {
	printData(args.binding.id + " has been selected. Click number: " + trackI);
	trackI++;
}

function logBindings() {
	Office.context.document.bindings.getAllAsync(function (asyncResult) {
		var bindingString = '';
		for (var i in asyncResult.value) {
			bindingString += asyncResult.value[i].id + '\n';
		}
		printData('Number of existing bindings: ' + asyncResult.value.length + '\n\nIds of existing bindings: \n' + bindingString);
	});
}

function fileName1() {
	printData(Office.context.document.url);
}

function fileName2() {
	printData(document.location);
}

function addCustomXML() {
	Office.context.document.customXmlParts.addAsync("<testns:book xmlns:testns='http://testns.com'><testns:page number='1'>Hello</testns:page><testns:page number='2'>world!</testns:page></testns:book>",
		function (asyncResult) {
			asyncResult.value.addHandlerAsync("nodeDeleted", function (asyncResult2) {
					printData(asyncResult2.type)
				},
				function (asyncResult3) {
					printData(asyncResult3.status)
				});
		});
}

function sendCommandToVB() {

	var macroName = jQuery('#sendData').val();
	var commandArray = "cmd[[[" + macroName + "]]]";
	var sXml = '<?xml version="1.0" encoding="UTF-8"?><inject>' + commandArray + '</inject>';

	Office.context.document.customXmlParts.addAsync(sXml, function (asyncResult) {
		var obj = asyncResult.value;
		printData("Add XML is done.");
		if (asyncResult && !asyncResult.error) {
			obj.deleteAsync(function (delRes) {
				printData("Delete is done.");
			});
		}
	});
}

function bindContentControl(alias, newID) {
	Office.context.document.bindings.addFromNamedItemAsync(alias, 'text', {
		id: newID
	}, function () {
		logBindings();
	});
}

/* magicWord is a hacky way to pass a message instead of the xmlListener; it fires a doc change event invisibly*/
/*
function magicWord() {
	var theID = jQuery('#sendData').val();
	var txt = '<span style="color:white;">mumbojumbo</span>';
	var coerc = 'html';

	Office.select("bindings#" + theID).setDataAsync(txt, {
		coercionType: coerc
	}, function (asyncResult) {
		var obj = asyncResult.value;
		printData("Add magicWord is done.");
		if (asyncResult && !asyncResult.error) {
			obj.deleteAsync(function (delRes) {
				printData("Delete is done.");
			});
		}
	});
}
*/

function goToByID() {
	var theID = $('#sendData').val();
	Office.context.document.goToByIdAsync(theID, "binding", function (result) {});
}

function WordToPaneXMLListener() {
    Office.context.document.customXmlParts.getByNamespaceAsync("http://testns.com", function (result) {
        
        result.value[0].addHandlerAsync(Office.EventType.nodeDeleted, nodeDeletedHandler, function(asyncResult){
            if (asyncResult.status === 'failed') {
                printData('NodeInserted Listener Failed to Add.\n\nError: ' + asyncResult.error.message);
            } else {
                printData('NodeInserted Listener Added.');
            }
        });
        
        function nodeDeletedHandler(eventArgs) {
            printData("A node has been inserted.");
        }
    });
}