import docxpy
import os
import random
from os import listdir
from os.path import isfile, join


class FreePlane():
    def printIndentationTabs(noOfTabs):
        tabContent = ""
        for i in range(noOfTabs):
            tabContent += '\t'
        return tabContent
    def cleanXMLfromSpecialChars(line):
        """
        Ampersand	&amp;	&
        Less-than	&lt;	<
        Greater-than	&gt;	>
        Quotes	&quot;	"
        Apostrophe	&apos;	'
        """
        return str(line).replace("&", "&amp;").replace("\"", "&quot;").replace("<", "&lt;").replace(">",
                                                                                                    "&gt;").replace("'",
                                                                                                                    "&apos;")
    def makeSelfContainedNode(text,link):
        if(link==""):
            node = '<node ID=\"ID_' + str(random.randint(1, 2000000000)) + '\"' + ' TEXT=\"' + text + "\"" + "/>\n"
        else:
            node = '<node ID=\"ID_' + str(random.randint(1, 2000000000)) + '\"' +' TEXT=\"' + text + "\""+ ' LINK=\"' + link  + "\"/>\n"
        return node
    def makeNode(text,link):
        if(link==""):
            node = '<node ID=\"ID_' + str(random.randint(1, 2000000000)) + '\"' + ' TEXT=\"' + text + "\"" + ">\n"
        else:
            node = '<node ID=\"ID_' + str(random.randint(1, 2000000000)) + '\"' +' TEXT=\"' + text + "\""+ ' LINK=\"' + link  + "\">\n"
        return node
    def createMindMapFile(mindmapContent,destFolder,fileName):
        completeFileName = destFolder + fileName
        outFile = open(completeFileName, 'w',encoding="utf-8")
        outFile.writelines(mindmapContent)
        return



def generateNodesFromDocHyperlinks(docHyperlinks,mindmap,noOfTabs,noOfNodesinNode):
    nodes = ""
    countofNodes = noOfNodesinNode
    firstFlag = True
    for hyperlink in docHyperlinks:
        if(countofNodes%noOfNodesinNode)==0:
            if(firstFlag):
                nodes = nodes + FreePlane.printIndentationTabs(noOfTabs)+FreePlane.makeNode("Node"+str(countofNodes/noOfNodesinNode),"")
                noOfTabs+=1
                firstFlag = False
            else:
                noOfTabs -= 1
                nodes = nodes + FreePlane.printIndentationTabs(noOfTabs) + "</node>\n"
                nodes = nodes + FreePlane.printIndentationTabs(noOfTabs) + FreePlane.makeNode("Node"+str(countofNodes/noOfNodesinNode),"")
                noOfTabs += 1
        nodes = nodes + FreePlane.printIndentationTabs(noOfTabs)+ FreePlane.makeSelfContainedNode(FreePlane.cleanXMLfromSpecialChars(str(hyperlink[0]).replace("b'","").replace("'","")),hyperlink[1])
        countofNodes += 1
    return nodes
def getDocFileName():
    onlyfiles = [f for f in listdir(str(os.curdir)) if isfile(join(str(os.curdir), f))]
    count_html = 0
    input_file_name = ""
    input_file_not_found_flag = True
    for file in onlyfiles:
        if '.docx' in file or '.doc' in file:
            count_html += 1
            input_file_not_found_flag = False
            input_file_name = file
    if count_html > 1:
        print("More than 1 input files found")
        a = input()
        return a
    if input_file_not_found_flag:
        print("No input File found")
        a = input()
        return a
    return input_file_name
def getLinksFromHTMLDoc():
    doc = docxpy.DOCReader(getDocFileName())
    doc.process()  # process file
    docHyperlinks = doc.data['links']
    noOfNodesinNode = 5
    mindMapName = "HTML page imported"
    mindmap = "<map version=\"1.0.1\">\n<node ID=\"ID_" + str(
        random.randint(1, 2000000000)) + "\" TEXT=\"" + mindMapName + "\">\n"
    noOfTabs = 1

    mindmap = mindmap + generateNodesFromDocHyperlinks(docHyperlinks, mindmap, noOfTabs, noOfNodesinNode)

    mindmap = mindmap + FreePlane.printIndentationTabs(noOfTabs) + "</node>\n</node>\n</map>"
    #print(mindmap)
    FreePlane.createMindMapFile(mindmap,"", "importedFromHTML.mm")
    return



getLinksFromHTMLDoc()
print("Links Imported.\n")
print("Press any key to continue.\n")
a = input()

