var text = "<p><b>Welcome <i>Ke</i></b><i>nn</i><b><i>y</i></b></p>";
var output = "";

var doc = new DOMParser().parseFromString(text, "text/html");
var pNodes = doc.getElementsByTagName("p");
console.log("Number of P nodes", pNodes.length, pNodes);

//Iterating the paragraph p nodes

for (var p = 0; p < pNodes.length; p++) {
    var pNode = pNodes[p];

    console.log("Start of node | <w:p>");
    output += "<w:p>";
    output += matchParagraphAlignment(pNode)

    var childNodes = pNode.childNodes;

    for (var i = 0; i < childNodes.length; i++) { //Navigating along the node
        var childNode = childNodes[i];
        console.log("Navigate Child Node > ", childNode.tagName, childNode.nodeValue);
        var innerTag = matchTagNameToNotation(childNode);
        if (childNode.nodeName == '#text') {
            output += "<w:r>";
            output += "<w:t xml:space=\"preserve\">" + childNode.textContent + "</w:t>";
            output += "</w:r>";
        }

        var grandchildNode = childNode.childNodes;
        for (var j = 0; j < grandchildNode.length; j++) {
            var gc = grandchildNode[j];
            output += printOutput(gc);
        }

    } //end of for-loop
    output += "</w:p>";
} //end of all p nodes

console.log(output);


function printOutput(node) {
    var output = '';
    //todo - check the length of the child node
    //the logic will be recurring for the grandchild node
    //lets assume there is not child node
    output += "<w:r>";
    if (node.nodeName == '#text') {
        if (innerTag != '') {
            output += "<w:rPr>";
            output += innerTag;
            output += "</w:rPr>";
        }
        output += "<w:t xml:space=\"preserve\">" + node.textContent + "</w:t>";
    }
    //If the node is not a text node, style has to apply
    else {
        output += "<w:rPr>";
        if (innerTag != '') {
            output += innerTag;
        }
        output += matchTagNameToNotation(node);
        output += "</w:rPr>";
        output += "<w:t xml:space=\"preserve\">" + node.textContent + "</w:t>";
    }

    output += "</w:r>";
    //console.log(node, output);
    return output;
}

function matchParagraphAlignment(node) {
    var output = '';
    if (node.style.textAlign != null && node.style.textAlign != '') {
        output += "<w:pPr>";
        output += "<w:jc w:val=\"" + node.style.textAlign + "\"/>";
        output += "</w:pPr>"
    }
    return output;
}

function matchTagNameToNotation(node) {
    var output = '';
    var tagName = node.tagName;
    console.log(tagName);
    if (tagName != null) {
        switch (tagName.toLowerCase()) {
            case 'i':
                output = "<w:i/>";
                break;
            case 'b':
                output = "<w:b/>";
                break;
            case 'u':
                output = "<w:u/>";
                break;
            case 'font':
                //Currently only color is supported in the RichText Dialog box
                var color = node.color;
                output = "<w:color  w:val=\"" + color + "\"/>"
            default:
                break;
        } //end switch
    }
    return output;
}