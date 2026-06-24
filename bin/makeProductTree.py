#!/usr/bin/env python
# Take list file with parents and make a tex diagram of the product tree.
# Top of the tree will be left of the page ..
# this allows a LONG list of products.
from __future__ import print_function

from treelib import Tree
import argparse
import csv
import re


import os
import os.path
import pickle

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from oauth2client.client import Credentials

import argparse

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Rubin Product Tree from Google Sheet'



# pt sizes for box + margin + gap between boex
txtheight = 35
leafHeight = 1.55  # cm space per leaf box .. height of page calc
leafWidth = 4.9  # cm space per leaf box .. width of page calc
smallGap = 0.2 # cm between leaf boxes in the same group
bigGap = 0.2 # cm between different levels, or leaf boxes
sep = 2  # inner sep
gap = 4
WBS = 1  # Put WBS on diagram
PKG = 1  # put packages on diagram
outdepth = 100  # set with --depth if you want a shallower tree

# Global team data for including in tree diagrams
g_team_data = None  # Dict {team_name: {institution: FTE, ...}, ...}
g_institutions = None  # List of institution names for ordering


def get_team_total_fte(orig_name):
    """Get total FTE for a team.
    
    Args:
        orig_name: Original (unescaped) product name to look up in team data
        
    Returns:
        Total FTE as float, or 0 if no data
    """
    if not g_team_data:
        return 0
    
    # Try to find matching team data (case-insensitive match)
    name_lower = orig_name.lower()
    for team_name, fte_dict in g_team_data.items():
        if team_name.lower() == name_lower:
            # Sum all institution FTEs
            total = sum(fte_dict.values())
            return total
    return 0


def get_team_fte_label(orig_name, prod_type=""):
    """Get formatted team FTE label for a product if team data is available.
    
    Args:
        orig_name: Original (unescaped) product name to look up in team data
        prod_type: Product type - only "Team" types get FTE labels
        
    Returns:
        String with institution FTEs formatted for LaTeX, or empty string if no data
    """
    if not g_team_data or not g_institutions:
        return ""
    
    # Only add team FTE info to Team boxes
    if prod_type.lower() != "team":
        return ""
    
    # Try to find matching team data (case-insensitive match)
    team_fte = None
    name_lower = orig_name.lower()
    for team_name, fte_dict in g_team_data.items():
        if team_name.lower() == name_lower:
            team_fte = fte_dict
            break
    
    if not team_fte:
        return ""
    
    # Format FTEs for display
    parts = []
    for inst in g_institutions:
        fte = team_fte.get(inst, 0)
        if fte > 0:
            parts.append(f"{inst}:{fte:.1f}")
    
    # Add 'Other' if present
    other = team_fte.get('Other', 0)
    if other > 0:
        parts.append(f"Oth:{other:.1f}")
    
    if parts:
        return r" \\ \scriptsize " + ", ".join(parts)
    return ""

def get_credentials() -> Credentials:
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def get_sheet(sheet_id, range):
    """
    grab the google sheet and return data from sheet
    :String sheetId: GoogelSheet Id like
        1R1h41KVtN2gKXJAVzd4KLlcF-FnNhpt1G06YhzwuWiY
    :String sheets: List of TabName!An:Zn  ranges

    """
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=range).execute()
    return result


class Product(object):
    def __init__(self, id, name, parent, desc, manager, owner, type, orig_name=None):
        self.id = id
        self.name = name
        self.orig_name = orig_name if orig_name else name  # Original unescaped name for team matching
        self.parent = parent
        self.desc = desc
        self.manager = manager
        self.owner = owner
        self.type = type


def constructTree(values, team_tree_mode=False):
    "scan the values file and construct  a tree structure"
    # assume root is after header on line 1
    skip_to = 2
    count = 0
    ptree = Tree()
    
    # In team tree mode, create a "top" root node first
    if team_tree_mode:
        top_prod = Product("top", "Rubin Observatory", "", "", "", "", "")
        ptree.create_node(top_prod.id, top_prod.id, data=top_prod)
        print("top is root (team tree mode)")
    
    for line in values:
        count = count + 1
        if count < skip_to:
            continue
        # Skip empty lines
        if not line or line[0].strip() == '':
            continue
        print(line)
        id = fixIdTex(line[0]) #make an id from the name
        pid= fixIdTex(line[1]) #use the same formaula on the parent name then we are good
        # In team tree mode, use column D (line[3]) for name; otherwise use column C (line[2])
        if team_tree_mode and len(line) > 3 and line[3].strip():
            orig_name = line[3].strip()  # Team name from column D for matching
            name = fixTex(line[3])
        else:
            orig_name = line[2].strip() if len(line) > 2 else ""
            name = fixTex(line[2])
        type = line[4]
        lead = "TBD"
        if (len(line) >= 6):
            lead = fixTex(line[5])
        po = ""
        if (len(line) >= 7):
            po = fixTex(line[6])
        # In team tree mode, column D is the team name, so notes is empty
        if team_tree_mode:
            notes = ""
        else:
            notes=fixTex(line[3])
            if len(line) == 8:
                notes= f"{notes}:{fixTex(line[7])}"
        prod = Product(id, name, pid, notes, lead, po, type, orig_name)
        if count == skip_to:  # first node from sheet
            if team_tree_mode:
                # In team tree mode, first node is child of "top"
                prod.parent = "top"
                ptree.create_node(prod.id, prod.id, data=prod, parent="top")
                print(f"{id} is child of top (team tree mode)")
            else:
                # Normal mode - first node is root
                print(f"{id} is root")
                ptree.create_node(prod.id, prod.id, data=prod)
        else:
            #print("Creating node:" + prod.id + " name:"+ prod.name +
            #      " parent:" + prod.parent)
            if prod.parent != "":
                ptree.create_node(prod.id, prod.id, data=prod,
                                  parent=prod.parent)
            elif team_tree_mode:
                # In team tree mode, nodes with no parent become children of "top"
                prod.parent = "top"
                ptree.create_node(prod.id, prod.id, data=prod, parent="top")
                print(f"{id} has no parent, assigned to top (team tree mode)")
            else:
                print(id + " no parent")

    print("{} Product lines".format(count))
    return ptree


def fixIdTex(text):
    id= re.sub(r"\s+","", text)
    id= id.replace("(","")
    id= id.replace(")","")
    id= id.replace("\"","")
    id= id.replace("_", "")
    id= id.replace(".","")
    id= id.replace("&","")

    return id

def fixTex(text):
    ret = text.replace("_", "\\_")
    ret = ret.replace("/", "/ ")
    ret = ret.replace("&", "\\& ")
    return ret


def slice(ptree, outdepth):
    if (ptree.depth() == outdepth):
        return ptree
    # copy the tree but stopping at given depth
    ntree = Tree()
    nodes = ptree.expand_tree()
    count = 0  # subtree in input
    for n in nodes:
        #print("Accesing {}".format(n))
        depth = ptree.depth(n)
        prod = ptree[n].data

        #print("outd={od} mydepth={d} Product: {p.id} name: {p.name} parent: {p.parent}".format(od=outdepth, d=depth, p=prod))
        if (depth <= outdepth):
            # print(" YES ", end='')
            if (count == 0):
                ntree.create_node(prod.id, prod.id, data=prod)
            else:
                ntree.create_node(prod.id, prod.id, data=prod,
                                  parent=prod.parent)
            count = count + 1
         #print()
    return ntree


def outputTexTable(tout, ptree):
    nodes = ptree.expand_tree()
    for n in nodes:
        prod = ptree[n].data
        print(r"{{\textbf{{{p.name}}}}} & {p.manager} & {p.owner} & {p.desc} \\ "
              r"\hline".format(p=prod), file=tout)
    return


def outputType(fout,prod):
    print("] {", file=fout, end='')
    print(r"\textbf{" + prod.name + "} ", file=fout, end='')
    team_label = get_team_fte_label(prod.orig_name, prod.type)
    if team_label:
        print(team_label, file=fout, end='')
    print("};", file=fout)
    if prod.type != "":
        # For Team types, add total FTE in parentheses
        if prod.type.lower() == "team" and g_team_data:
            total_fte = get_team_total_fte(prod.orig_name)
            if total_fte > 0:
                type_label = f"{prod.type} ({total_fte:.1f})"
            else:
                type_label = prod.type
            print(r"\node [below right] at ({p.id}.north west) {{\small \color{{blue}}{label}}} ;".format(p=prod, label=type_label), file=fout)
        else:
            print(r"\node [below right] at ({p.id}.north west) {{\small \color{{blue}}{p.type}}} ;".format(p=prod), file=fout)
    return

def parent(node):
    return node.data.parent
def id(node):
    return node.data.id

def groupByParent(nodes):
    nn = []
    pars = []
    for n in nodes:
        parent = n.data.parent
        if parent not in pars:
            pars.append(parent)
            for n2 in nodes:
                if (n2.data.parent == parent) :
                   #print (r"Got {n.data.id} par {n.data.parent} looking for {par}".format(n=n2,par=parent))
                   nn.append(n2)
    return nn;

def organiseRow(r, rowMap): #group according to parent in order of prevous row
    prow=rowMap[r-1]
    row = rowMap[r]
    print(r" organise {n} nodes in  Row {r} by {nn} nodes in {r}-1".format(r=r, n=len(row),nn=len(prow)))
    nrow= []
    for p in prow: #scan parents
        parent = p.data.id
        for n2 in row:
           if (n2.data.parent == parent) :
                   print (r"Got {n.data.id} par {n.data.parent} looking for {par}".format(n=n2,par=parent))
                   nrow.append(n2)
    rowMap[r]=nrow
    return

def drawLines(fout,row):
    #print(r"   nodes={n} - now lines for row={n} nodes  ".format( n=len(row)))
    for p in row:
        prod=p.data
        print(r" \draw[pline]   ({p.id}.north) -- ++(0.0,0.5) -| ({p.parent}.south) ; ".format(p=prod),
             file=fout )
    return

def layoutRows(fout, rowMap, start, end, count, ptree, children,childcount, goingDown ):
    prow=None
    inc = -1
    if goingDown==1:
        inc=+1
    for r in range(start,end,inc): # Printing last row first
        row = rowMap[r]
        count = count + doRow(fout,ptree,children,row,r,childcount, goingDown)
        print(r"Output  depth={d},   nodes={n} start={s} end={e} goingDown={a}".format(d=r, n=len(row), s=start, e=end, a=goingDown))
        if (goingDown==1): #draw lines between current row and parents
            drawLines(fout,row)
        if (prow and goingDown == 0):
           # print(r"Output  depth={d},   nodes={n} - now lines for prow={pr} nodes".format(d=r, n=len(row), pr=len(prow)))
            drawLines(fout,prow)
            prow=row
            row=[]
        else:
            prow=row
    return count

def outputLandMix(fout,ptree):
# attempt to put DM on top of page then each of the top level sub trees in portrait beneath it accross the page ..
    stub = slice(ptree, 1)
    nodes = stub.expand_tree(mode=Tree.WIDTH) # default mode=DEPTH
    row = []
    count =0
    root= None
    child = None
    for n in nodes:
        count= count +1
        if (count ==1): #  root node
           root=ptree[n].data
        else:
           row.append(ptree[n])

    child = row[count//2].data
    sib=None
    count =1 # will output root after
    prev= None
    for n in row: # for each top level element put it out in portrait
        p = n.data
        stree = ptree.subtree(p.id)
        d = 1
        if (prev):
            d = prev.depth()
        width =  d  * ( leafWidth + bigGap ) + bigGap # cm
        if sib:
            print(sib.name, d, p.name)
        print(r" {p.id} {p.parent} depth={d} width={w} ".format(p=p, d=d,w=width ))
        count = count + outputTexTreeP(fout, stree, width, sib, 0)
        sib=p
        prev = stree

    ### place root node
    team_label = get_team_fte_label(root.orig_name, root.type)
    print(r"\node ({p.id}) "
         r"[wbbox, above=15mm of {c.id}]{{\textbf{{{p.name}}}{team_label}}};".format(p=root,c=child,team_label=team_label),
         file=fout)
    drawLines(fout,row)
    print("{} Product lines in TeX ".format(count))
    return

def land_red_leaves2(ptree, debug):
    # returns the span of a subtree,
    d = ptree.depth()
    stub = slice(ptree, 1)
    if debug:
        print(stub)
    if d==0:
        if debug:
            print('  --  no depth', ptree)
        return(1)
    nodes = stub.expand_tree(mode=Tree.WIDTH)
    row = []
    count = 0
    root = None
    nl = 0
    prevw = None
    pnl = 0
    pd = 0
    plad = 0
    alpd = 0
    for n in nodes:
        count = count +1
        #print('node', n, str(count))
        if (count ==1): #  root node, nothing to do here
           root=ptree[n].data
           if debug:
                print('----------', root.id)
        else:
            stree = ptree.subtree(n)
            p = ptree[n].data
            sdepth = stree.depth()
            if prevw:
                if (sdepth == 0):
                    al = 1
                    plad = 1
                    #delta = 1
                elif (sdepth<pd):
                    sprevw = slice(prevw, sdepth)
                    al = land_red_leaves2(stree, debug)
                    plad = land_red_leaves2(sprevw, debug)
                    #delta = plad / 2 + al - pnl / 2
                elif (sdepth>pd):
                    astree = slice(stree, pd)
                    alpd = land_red_leaves2(astree, debug)
                    al = land_red_leaves2(stree, debug)
                    #delta = (alpd + al) /2
                else:
                    al = land_red_leaves2(stree, debug)
                    #delta = al
                if (sdepth<pd):
                    delta = plad / 2 + al - pnl / 2
                elif (sdepth>pd):
                    delta = (alpd + al) /2
                else:
                    delta = al
                nl = nl + delta
                if debug:
                    print(p.id, sdepth, pd, al, pnl, plad, alpd, delta, nl)
                pnl = al
            else:
                if (sdepth == 0):
                    nl = 1
                else:
                    nl = land_red_leaves2(stree, debug)
                pnl = nl
            prevw = stree
            pd = sdepth
    if debug:
        print('RETURN', root.id, nl)
    return(nl)

def land_red_leaves(ptree):
    # returns the span of a subtree,
    d = ptree.depth()
    stub = slice(ptree, 1)
    #print('STUB\n', stub)
    if d==0:
        #print('  --  no depth', ptree)
        return(1)
    nodes = stub.expand_tree(mode=Tree.WIDTH)
    row = []
    count = 0
    root = None
    nl = 0
    prevw = None
    pnl = 0
    pd = 0
    for n in nodes:
        count = count +1
        #print('node', n, str(count))
        if (count ==1): #  root node, nothing to do here
           root=ptree[n].data
           #print('----------', root.id)
        else:
            stree = ptree.subtree(n)
            p = ptree[n].data
            sdepth = stree.depth()
            if (prevw):
                sprevw = slice(prevw, sdepth)
                if sdepth == 0:
                    al = 1
                    plad = 1
                else:
                    #print(' call P')
                    plad = land_red_leaves(sprevw)
                    #print(' call 1', p.id)
                    al = land_red_leaves(stree)
                #print(' compare', p.id, sdepth, tmp, pnl, pd)
                #if (sdepth == 0 and pd == 1 and tmp == 1 and pnl > 1):
                #    nl = nl
                #else:
                #    nl = nl + tmp
                if sdepth < pd:
                    delta = al - (pnl - plad)
                    #print(' compare', p.id, sdepth, pd, al, pnl, pd, delta)
                else:
                    delta = al
                nl = nl + delta
                pnl = al
                #print(' - node', p.id, nl)
            else:
                if sdepth==0:
                    nl = 1
                else:
                    #print(' call 0', p.id)
                    nl = land_red_leaves(stree)
                pnl = nl
                #print(' - primeN', p.id, nl)
            prevw = stree
            #pnl = land_red_leaves(stree)
            pd = sdepth
    #print('RETURN', root.id, nl)
    return(nl)

def outputLandR(fout, ptree, pid):
    stub = slice(ptree, 1) # I want to print only one level down
    nodes = stub.expand_tree(mode=Tree.WIDTH)
    row = []
    count =0
    root= None
    child = None
    for n in nodes:
        count= count +1
        if (count ==1): #  root node
           root=ptree[n].data
        else:
           row.append(ptree[n])

    child = row[(count-1)//2].data # the mid child, the parent will be positioned above it
    sib=None
    count =1 # the root
    prev= None
    prevw = None
    pnl = None
    tnnl = []
    for n in row:
        prod = n.data
        stree = ptree.subtree(prod.id)
        sdepth = stree.depth()
        if (sib):
            sprevw = slice(prevw, sdepth) # Slice previus subtree that can acomodated with the actual one
            nleaves = len(sprevw.leaves())
            #nleaves = land_red_leaves(sprevw)
            btype='p'
            if prod.type and len(prod.type)>0:
                btype=prod.type[0].lower()
            #    print(nleaves, prod.name)
            #print('Previous tree:\n', sprevw)
            nl = land_red_leaves(sprevw)
            nleaves = nl
            dist = (nleaves -1) * 109 + gap
            team_label = get_team_fte_label(prod.orig_name, prod.type)
            print(rf"\node ({prod.id}) "
                 rf"[{btype}box, right={dist}pt of {sib.id}]{{\textbf{{{prod.name}}}{team_label}}};",file=fout)
        else:
            team_label = get_team_fte_label(prod.orig_name, prod.type)
            if (pid):
                print(rf"\node ({prod.id}) "
                     rf"[{btype}box, below=15mm of {pid}]{{\textbf{{{prod.name}}}{team_label}}};",
                     file=fout)
            else:
                print(fr"\node ({prod.id}) "
                     fr"[{btype}box]{{\textbf{{{prod.name}}}{team_label}}};", file=fout)
        sib = prod
        prevw = stree
        if (sdepth > 0):
            outputLandR(fout, stree, prod.id)

    if (pid):
        drawLines(fout,row)
    else:
        ### place root node
        team_label = get_team_fte_label(root.orig_name, root.type)
        print(r"\node ({p.id}) "
             r"[wbbox, above=15mm of {c.id}]{{\textbf{{{p.name}}}{team_label}}};".format(p=root,c=child,team_label=team_label),
             file=fout)
        print("{} Product lines in TeX ".format(count))
        drawLines(fout,row)
    return

# fout: output file
# ptree: the tree to print out
# pid: parent id (seems superfluous since it coincides with root.id)
# prevl: previous tree number of leaves
def outputLandR2(fout, ptree, pid, prevd, prevl):
    stub = slice(ptree, 1) # I want to print only one level down
    nodes = stub.expand_tree(mode=Tree.WIDTH)
    row = []
    count =0
    root= None
    child = None
    for n in nodes:
        count= count +1
        if (count ==1): #  root node
           root=ptree[n].data
        else:
           row.append(ptree[n])

    child = row[(count-1)//2].data # the mid child, the parent will be positioned above it
    nch = count - 1
    #print('root', root.id, nch)
    sib=None
    count =1 # the root
    prevw = None
    pnl = 0
    pdph = 0
    plad = 0
    alpd = 0
    for n in row:
        prod = n.data
        stree = ptree.subtree(prod.id)
        sdepth = stree.depth()
        #if prod.id == 'appipe':
        #    print(stree)
        #    print('reduced leaves', prod.id, nl)
        #nl = len(stree.leaves())
        if (sib):
            if (sdepth == 0):
                al = 1
                plad = 1
            elif (sdepth < pdph):
                sprevw = slice(prevw, sdepth) # Slice previus subtree that can collide with the actual one
                if sdepth==1:
                    plad = 1
                else:
                    #if prod.id=='lsstobs':
                    #    plad = land_red_leaves2(sprevw, 'D')
                    #else:
                    plad = land_red_leaves2(sprevw, None)
                al = land_red_leaves2(stree, None) #I get the number of leaves of the subtree
            elif (sdepth > pdph):
                astree = slice(stree, pdph)
                if pdph==1:
                    alpd = 1
                else:
                    #if prod.id=='jointcal':
                    #    alpd = land_red_leaves2(astree, 'debug')
                    #else:
                    alpd = land_red_leaves2(astree, None)
                al = land_red_leaves2(stree, None)
            else:
                al = land_red_leaves2(stree, None)

            if (sdepth<pdph):
                delta = al / 2 + plad / 2
            elif (sdepth>pdph):
                delta = alpd / 2 + pnl / 2
            else:
                delta = al / 2 + pnl / 2
            #delta = nl / 2 + nleaves / 2
            dist = (delta -1) * 109 + gap * (delta + 1)
            print(r"Inspect Line: prod: {p} - al: {al} - pnl: {pnl} - plad: {plad} - alpd: {alpd} - sdepth: {sd} - pd: {pd} - delta: {delta} - dist: {dist}".format(p=prod.id,al=al,pnl=pnl,plad=plad,sd=sdepth,pd=pdph,delta=delta,dist=dist,alpd=alpd))
            team_label = get_team_fte_label(prod.orig_name, prod.type)
            print(r"\node ({p.id}) "
                 r"[pbox, right={d}pt of {s.id}]{{\textbf{{{p.name}}}{team_label}}};".format(p=prod,s=sib,d=dist,team_label=team_label),
                 file=fout)
        else:
            #
            al = land_red_leaves2(stree, None)
            team_label = get_team_fte_label(prod.orig_name, prod.type)
            if (pid):
                dist = 109 * ( (nch - 1 ) / 2 - 1 ) + gap * ( ( nch - 1 ) / 2 + 1 )
                #dist = 109 * ( (al - 1 ) / 2 - 1 ) + gap * ( ( al - 1 ) / 2 + 1 )
                #print(prod.id, pid, nch, dist)
                print(r"\node ({p.id}) "
                     r"[pbox, below left=15mm and {d}pt of {pid}]{{\textbf{{{p.name}}}{team_label}}};".format(p=prod,pid=pid,d=dist,team_label=team_label),
                     file=fout)
            else:
                print(r"\node ({p.id}) "
                     r"[pbox]{{\textbf{{{p.name}}}{team_label}}};".format(p=prod,team_label=team_label),
                     file=fout)
        if (sdepth > 0):
            outputLandR2(fout, stree, prod.id, pdph, pnl)
        sib = prod
        if (pnl < al and sdepth>=pdph):
            prevw = stree
            pnl = al
            pdph = sdepth

    if (pid):
        drawLines(fout,row)
    else:
        ### place root node
        team_label = get_team_fte_label(root.orig_name, root.type)
        print(r"\node ({p.id}) "
             r"[wbbox, above=15mm of {c.id}]{{\textbf{{{p.name}}}{team_label}}};".format(p=root,c=child,team_label=team_label),
             file=fout)
        print("{} Product lines in TeX ".format(count))
        drawLines(fout,row)
    return


def outputLandW(fout,ptree):
    childcount = dict() # map of counts of children
    children= dict() # map of most central child to place this node ABOVE it
    rowMap = dict()
    nodes = ptree.expand_tree(mode=Tree.WIDTH) # default mode=DEPTH
    count = 0
    row=[]
    depth=ptree.depth()
    d=0
    pdepth=d
    prow= None
    pn= None
    cc =0
    # first make rows
    for n in nodes:
        count= count +1
        prod =ptree[n].data
        if (not pn):
            pn=prod
        d = ptree.depth(prod.id)
        # count the children as well
        if ( not pn.parent == prod.parent):
            childcount[pn.parent]=cc
            print(r" Set {p.parent} : {cc} children".format(p=pn, cc=cc ))
            cc=0
        cc= cc+1
        if d != pdepth: # new row
            #print(r" depth={d},   nodes={n}".format(d=pdepth, n=len(row)))
            rowMap[pdepth] = row
            row=[]
            pdepth=d
            pn= None
        row.append(ptree[n])
        pn = prod
    rowMap[d] = row # should be root

    childcount[pn.parent]=cc
    print(r"Out of loop  depth={d}, rows={r}  nodes={n}".format(d=depth, r=len(rowMap), n=count))
    count=0
    # now group the children under parent .. should be done by WIDT FIRST walk fo tree
    # for r in range(2,depth,1): # root is ok and the next row
    #   organiseRow(r,rowMap)
    #now actually make the tex
    # need to  find row with most leaves .. then layout relative to that..
    wideR = depth
    for r in range(depth,-1,-1): # Look at each row
        rowSize = len(rowMap[r])
        if rowSize > len(rowMap[wideR]):
            wideR=r
    print(r"Widest row  depth={d},   nodes={n} layout {d} to  -1".format(d=wideR, n=len(rowMap[wideR])))
    #now lay out row wideR and UP to root last 0 indicated goingUpward
    count = count + layoutRows(fout,rowMap, wideR, -1, count, ptree, children, childcount, 0 )
    if (wideR != depth):
        print(r"Layout remainder down wideR={w} depth={d}".format(w=wideR+1, d=depth))
        # and layout the the widest row to the bottowm downward ,.
        count = count + layoutRows(fout,rowMap, wideR+1, depth+1, count, ptree, children,childcount,1 )

    print("{} Product lines in TeX ".format(count))
    return

def doRow(fout,ptree,children,nodes,depth, childcount, goingDown):
#Assuming the nodes are sorted by parent .. putput the groups of siblings and record
# children the middle child of each group
# this is for landscaepe outout but gets too wide wut full tree
    sdist=15  #mm  sibling group distance for equal distribution
    ccount=0;
    prev = Product("n", "n", "n", "n", "n", "n", "n")
    sibs = []
    child = None
    ncount= len(nodes)
    pushd=0
    for n in nodes:
        placed=0
        prod = n.data
        ccount = ccount + 1
        if (prod.id in children):
           child = children[prod.id]
        else:
           child= None
        if (depth==0):  # root node
           #print(r"depth==0 {p.id}  parent  {p.parent},   child={c}".format(p=prod, c=child))
           team_label = get_team_fte_label(prod.orig_name, prod.type)
           print(r"\node ({p.id}) "
               r"[wbbox, above=15mm of {c}]{{\textbf{{{p.name}}}{team_label}}};".format(p=prod,c=child,team_label=team_label),
               file=fout)
           placed=1
        else:
           print(r"\node ({p.id}) [pbox,".format(p=prod), file=fout, end='')
           if child and goingDown==0: # easy case - node aboove child
              print("above={d}mm of {c}".format(d=sdist,c=child), file=fout, end='')
              placed=1
           if goingDown==1  and not prev.parent == prod.parent:
              if not ccount==1: # if its the first one just put it left
                 ddist=sdist
                 if childcount[prod.parent] > 4 : # I got siblings
                     if pushd==0:
                         pushd=1
                         ddist= 3* sdist
                     else:
                         pushd=0
                 print("below={d}mm of {p.parent}".format(d=ddist,p=prod), file=fout, end='')
              placed=1
           # need to deal with next children
           dist=1 # siblings close then gap
           if ((prev.parent != prod.parent and ccount >1) or (ccount==ncount) ): # not forgetting the last group
              if (ccount==ncount): #we ar eon the last group tha tis the one we do not prev
                theProd=prod
              else:
                theProd=prev
                dist=sdist
              sibs = ptree.siblings(theProd.id)
              sc = len (sibs)
              msib= (int) ((float) (sc) / 2.0 )
              if (msib !=0):
                  children[theProd.parent] = sibs[msib].data.id
                  #print(r" parent  {pr.parent} over  prod {p.id}".format(pr=theProd, p=sibs[msib].data))
              else: #only child
                  children[theProd.parent] = theProd.id
                  #print(r" Only child or 1 sibling.  parent  {p.parent} over  prod {p.id} nsibs={sc}".format(p=theProd, sc=sc))
              sibs = []
           if (not ccount==1 and not child  and goingDown==0 or (goingDown ==1 and placed==0 )): # easy put out to right
                # distance should account for how many children lie beneath the sibling to the left
              if prev and prev.id in childcount:
                 dist = childcount[prev.id] * 15  + 1
                 #print(r" dist  {d} prev {p.id} {nc} children".format(p=prev, d=dist, nc=childcount[prev.id]))
              print("right={d}mm of {p.id}".format(d=dist,p=prev), file=fout, end='')
              placed=1
           #print(r"mydepth={md} depth={dp} {p.id} right={d}mm parent={p.parent} prevparent={pr.parent}"
           #         " prev={pr.id}".format(md=ptree.depth(prod.id),dp=depth,p=prod,d=dist, pr=prev))
           outputType(fout,prod)
           prev = prod
    return ccount


def outputTexTree(fout, ptree, paperwidth):
    count = outputTexTreeP(fout, ptree, paperwidth, None, 1)
    print("{} Product lines in TeX ".format(count))
    return

def outputTexTreeP(fout, ptree, width, sib, full):
    fnodes = []
    nodes = ptree.expand_tree() # default mode=DEPTH
    count = 0
    prev = Product("n", "n", "n", "n", "n", "n", "n")
    nodec =1
    # Text height + the gap added to each one
    blocksize = txtheight + gap + sep
    for n in nodes:
        prod = ptree[n].data
        fnodes.append(prod)
        depth = ptree.depth(n)
        count = count + 1
        # print("{} Product: {p.id} name: {p.name}"
        #       " parent: {p.parent}".format(depth, p=prod))
        bcode = 'p'
        if prod.type and len(prod.type) > 1:
            bcode = prod.type[0].lower()
        if (depth <= outdepth):
            if (count == 1 ):  # root node
                if full ==1:
                   team_label = get_team_fte_label(prod.orig_name, prod.type)
                   print(r"\node ({p.id}) "
                      r"[wbbox]{{\textbf{{{p.name}}}{team_label}}};".format(p=prod,team_label=team_label),
                      file=fout)
                else: #some sub tree
                   print(fr"\node ({prod.id}) [{bcode}box, ", file=fout)
                   if (sib):
                      print(f"right={width}cm of {sib.id}", file=fout, end='')
                   outputType(fout,prod)

            else:
                print(fr"\node ({prod.id}) [{bcode}box,", file=fout, end='')
                if (prev.parent != prod.parent):  # first child to the right if portrait left if landscape
                    found = 0
                    scount = count - 1
                    while found == 0 and scount > 0:
                        scount = scount - 1
                        found = fnodes[scount].parent == prod.parent
                    if scount <= 0:  # first sib can go righ of parent
                        print("right=15mm of {p.parent}".format(p=prod),
                              file=fout, end='')
                    else:  # Figure how low to go  - find my prior sibling
                        psib = fnodes[scount]
                        leaves = ptree.leaves(psib.id)
                        depth = len(leaves) - 1
                        lleaf = leaves[depth-1].data
                        # print("Prev: {} psib: {} "
                        #       "llead.parent: {}".format(prev.id, psib.id,
                        #                                 lleaf.parent))
                        ##if (lleaf.parent == psib.id):
                        ##    depth = depth - 1
                        # if (prod.id=="L2"):
                        #     depth=depth + 1 # Not sure why this is one short
                        # the number of leaves below my sibling
                        dist = depth * blocksize + gap
                        # print("{p.id} Depth: {} dist: {} blocksize: {}"
                        #       " siblin: {s.id}".format(depth, dist,
                        #                                s=psib, p=prod))
                        print("below={}pt of {}".format(dist, psib.id),
                              file=fout, end='')
                else:
                    # benetih the sibling
                    dist = gap
                    print("below={}pt of {}".format(dist, prev.id), file=fout, end='')
                outputType(fout,prod)
                print(r" \draw[pline] ({p.parent}.east) -| ++(0.4,0) |- ({p.id}.west); ".format(p=prod), file=fout)
            prev = prod
    return count


def mixTreeDim(ptree):
    "Return the max number of elements (hight and width) in a mixed tree."

    stub = slice(ptree, 1) # I want to print only one level down
    nodes = stub.expand_tree(mode=Tree.WIDTH)
    row = []
    count =0
    root= None
    n2l = 0
    nmaxSub = 0
    for n in nodes:
        count= count +1
        if (count == 1): #  root node
           root=ptree[n].data
        else:
            prod = ptree[n].data
            stree = ptree.subtree(prod.id)
            sdepth = stree.depth()
            n2l = n2l + 1 + sdepth
            subL = len(ptree.leaves(n))
            if subL > nmaxSub:
               nmaxSub = subL

    #nodes = ptree.expand_tree()
    #for n in nodes:
    #    depth = ptree.depth(n)
    #    #print(depth)
    #    if depth == 1:
    return (n2l, nmaxSub)

def makeTree(values, team_data=None, institutions=None):
    """This processes the google sheet produces a tex tree diagram and a tex longtable.
    
    Args:
        values: List of rows from the Google Sheet
        team_data: Optional dict {team_name: {institution: FTE_total, ...}, ...}
        institutions: Optional list of institution names for ordering
    """
    global g_team_data, g_institutions
    g_team_data = team_data
    g_institutions = institutions
    
    # Use TeamTree filename when team data is included
    base_name = "TeamTree" if team_data else "ProductTree"
    nf = f"{base_name}.tex"
    if (land!=None):
       nf = f"{base_name}Land.tex"
    print('Saving product tree in: ', nf)
    nt = "productlist.tex"

    # need to skip a line or two
    team_tree_mode = team_data is not None
    ptree = constructTree(values, team_tree_mode)

    paperwidth = 0
    height = 0
    if (outdepth <= 100 ):
        ntree = slice(ptree, outdepth)
    else:
        ntree = ptree
        #if (land!=1):
        #    paperwidth = 2
        #    height = -3

    n2, nMS = mixTreeDim(ntree)

    #print('>n2 - tree depth: ', n2, ntree.depth(), nMS)

    # Adjust box and page sizes for team mode
    global txtheight, leafHeight, leafWidth
    if team_tree_mode:
        txtheight = 56  # Larger boxes for team info
        leafHeight = 2.8  # Larger spacing for page height calc
        leafWidth = 5.9  # Larger spacing for page width calc
    else:
        txtheight = 35  # Default size
        leafHeight = 1.55  # Default spacing
        leafWidth = 4.9  # Default width

    # ptree.show(data_property="name")
    if (land==1):   #full landscape
      # get the number of groups of leaves
      tree_depth = ntree.depth()
      reduced_tree = slice(ntree, tree_depth -1)
      paperwidth = paperwidth + len(ntree.leaves()) * ( leafWidth + smallGap ) + len(reduced_tree.leaves()) * bigGap # cm
      height = height + ( ntree.depth() + 1 ) * ( leafHeight + 1.5 )  # cm
    elif (land==2):  #mixed landscape/portrait
      paperwidth = paperwidth + 5.2 * n2 + 0.7 # cm
      height = height + nMS * leafHeight #1.6  # cm
    elif (land==3):  #recursive landscape, same spacing
      # get the number of groups of leaves
      #print('-------------------------')
      nl = land_red_leaves(ntree)
      print('-------------------------', nl)
      paperwidth = paperwidth + nl  * ( leafWidth + smallGap ) # cm
      height = height + ( ntree.depth() + 1 ) * ( leafHeight + 1.5 )  # cm
    elif (land==0):
      nl = land_red_leaves2(ntree, None)
      print('-------------------------', nl)
      paperwidth = paperwidth + nl  * ( leafWidth + smallGap ) # cm
      height = height + ( ntree.depth() + 1 ) * ( leafHeight + 1.5 )  # cm
    else:
      paperwidth = paperwidth + ( ntree.depth() + 1 ) * ( leafWidth + bigGap ) # cm
      streew=paperwidth
      height = len(ntree.leaves()) * leafHeight + 0.5 # cm

    print('height:', height, '; width:', paperwidth)

    with open(nf, 'w') as fout:
        header(fout, paperwidth, height)
        if (land==0):
            outputLandR2(fout, ntree, None, None, None)
        elif (land==1):
            outputLandW(fout, ntree)
        elif (land==2):
            outputLandMix(fout, ntree)
        elif (land==3):
            outputLandR(fout, ntree, None)
        else:
            outputTexTree(fout, ntree, paperwidth)
        footer(fout)

    # Output team CSV in team mode, productlist.tex otherwise
    if team_tree_mode:
        output_team_csv(ptree, "teamlist.csv")
    else:
        with open(nt, 'w') as tout:
            theader(tout)
            outputTexTable(tout, ptree)
            tfooter(tout)

    return
# End makeTree


def output_team_csv(ptree, filename):
    """Output team data as CSV with institution FTE columns.
    
    Args:
        ptree: The product tree
        filename: Output CSV filename
    """
    if not g_team_data or not g_institutions:
        print("No team data available for CSV output")
        return
    
    print(f"Saving team list in: {filename}")
    
    with open(filename, 'w', newline='') as csvfile:
        # Build header: ID, Name, Parent, Type, Manager, Owner, then institution columns, then Total
        headers = ['ID', 'Name', 'Parent', 'Type', 'Manager', 'Owner'] + list(g_institutions) + ['Other', 'Total']
        
        # Write header
        csvfile.write(','.join(headers) + '\n')
        
        # Iterate through tree nodes
        nodes = ptree.expand_tree()
        for n in nodes:
            prod = ptree[n].data
            
            # Get team FTE data if available
            team_fte = None
            if prod.orig_name:
                name_lower = prod.orig_name.lower()
                for team_name, fte_dict in g_team_data.items():
                    if team_name.lower() == name_lower:
                        team_fte = fte_dict
                        break
            
            # Build row
            row = [
                prod.id,
                prod.orig_name or prod.name,
                prod.parent,
                prod.type,
                prod.manager.replace(',', ';'),  # Escape commas
                prod.owner.replace(',', ';'),
            ]
            
            # Add institution FTEs
            total = 0
            for inst in g_institutions:
                fte = team_fte.get(inst, 0) if team_fte else 0
                row.append(f"{fte:.1f}" if fte > 0 else "0")
                total += fte
            
            # Add Other and Total
            other = team_fte.get('Other', 0) if team_fte else 0
            total += other
            row.append(f"{other:.1f}" if other > 0 else "0")
            row.append(f"{total:.1f}" if total > 0 else "0")
            
            csvfile.write(','.join(row) + '\n')
    
    print(f"Team CSV written with {len(list(ptree.expand_tree()))} entries")


def theader(tout):
    print("""
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%  Product table generated by {} do not modify.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
""".format(__file__), file=tout)
    print(r"""\scriptsize
\begin{longtable} {
|p{0.2\textwidth}   |p{0.15\textwidth}|p{0.15\textwidth} |p{0.4\textwidth}|}
\multicolumn{1}{c|}{\textbf{Product}} &
\multicolumn{1}{c|}{\textbf{Manager}} &
\multicolumn{1}{c|}{\textbf{Owner}} &
\multicolumn{1}{c}{\textbf{Notes}}|\\ \hline""",
          file=tout)

    return


def header(fout, pwidth, pheight):
    print(r"""%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%
% Document:      DM  operations product tree
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
\documentclass{article}
\usepackage{times,layouts}
\usepackage{tikz,hyperref,amsmath}
\usetikzlibrary{positioning,arrows,shapes,decorations.shapes,shapes.arrows}
\usetikzlibrary{backgrounds,calc}""", file=fout)
    print(r"\usepackage[paperwidth={}cm,paperheight={}cm,".format(pwidth,
          pheight), file=fout)
    print(r"""left=-2mm,top=3mm,bottom=0mm,right=0mm,
noheadfoot,marginparwidth=0pt,includemp=false,
textwidth=30cm,textheight=50mm]{geometry}
\newcommand\showpage{%
\setlayoutscale{0.5}\setlabelfont{\tiny}\printheadingsfalse\printparametersfalse
\currentpage\pagedesign}
\hypersetup{pdftitle={DM products }, pdfsubject={Diagram illustrating the
products in LSST DM }, pdfauthor={ William O'Mullane}}
\tikzstyle{tbox}=[rectangle,text centered, text width=30mm]
\tikzstyle{wbbox}=[rectangle, rounded corners=3pt, draw=black, top color=blue!50!white, bottom color=white, very thick, minimum height=12mm, inner sep=2pt, text centered, text width=30mm]""", file=fout)

    print(r"\tikzstyle{dbox}=[rectangle, rounded corners=3pt, draw=black, top"
          " color=green!50!white, bottom color=white, very thick,"
          " minimum height=" + str(txtheight) + "pt, inner sep=" + str(sep) +
          "pt, text centered, text width=35mm]", file=fout)

    print(r"\tikzstyle{pbox}=[rectangle, rounded corners=3pt, draw=black, top"
          " color=yellow!50!white, bottom color=white, very thick,"
          " minimum height=" + str(txtheight) + "pt, inner sep=" + str(sep) +
          "pt, text centered, text width=35mm]", file=fout)

    print(r"\tikzstyle{abox}=[rectangle, rounded corners=3pt, draw=black, top"
          " color=green!50!white, bottom color=white, very thick,"
          " minimum height=" + str(txtheight) + "pt, inner sep=" + str(sep) +
          "pt, text centered, text width=35mm]", file=fout)

    print(r"\tikzstyle{tbox}=[rectangle, rounded corners=3pt, draw=black, top"
          " color=blue!50!white, bottom color=white, very thick,"
          " minimum height=" + str(txtheight) + "pt, inner sep=" + str(sep) +
          "pt, text centered, text width=44mm]", file=fout)

    print(r"\tikzstyle{gbox}=[rectangle, rounded corners=3pt, draw=black, top"
          " color=cyan!50!white, bottom color=white, very thick,"
          " minimum height=" + str(txtheight) + "pt, inner sep=" + str(sep) +
          "pt, text centered, text width=35mm]", file=fout)

    print(r"""\tikzstyle{pline}=[-, thick]
\begin{document}
\begin{tikzpicture}[node distance=0mm]""", file=fout)

    return


def footer(fout):
    print(r"""\end{tikzpicture}
\end{document}""", file=fout)
    return


def tfooter(tout):
    print(r"""\end{longtable}
\normalsize""", file=tout)
    return




DEFAULT_INSTITUTIONS = ['SLAC', 'IN2P3', 'UK', 'AURA', 'UW', 'Princeton']


def process_team_sheet(values, institutions):
    """
    Process team sheet data and calculate FTE totals per team by institution.

    Handles pivot table format where department/team names only appear on first row,
    and subsequent rows for same team have empty department/team columns.

    Args:
        values: List of rows from the Google Sheet
        institutions: List of institution names to track individually (others become 'Other')

    Returns:
        tuple: (team_data, dept_teams)
            team_data: {team_name: {institution: FTE_total, ...}, ...}
            dept_teams: {department: [team_name, ...], ...} - preserves order
    """
    team_data = {}
    dept_teams = {}  # Maps department to list of teams (in order)
    skip_to = 2  # Skip header row
    count = 0
    current_dept = None  # Track current department for pivot table format
    current_team = None  # Track current team for pivot table format

    # Create case-insensitive lookup for institutions
    inst_lookup = {inst.lower(): inst for inst in institutions}

    for row in values:
        count += 1
        if count < skip_to:
            continue

        # Column A (index 0): department
        # Column B (index 1): team name
        # Column C (index 2): institution
        # Column D (index 3): FTE
        if len(row) < 4:
            continue

        dept_cell = row[0].strip() if len(row) > 0 and row[0] else ""
        team_name_cell = row[1].strip() if len(row) > 1 and row[1] else ""
        institution_raw = row[2].strip() if len(row) > 2 and row[2] else ""
        fte_str = row[3].strip() if len(row) > 3 and row[3] else "0"

        # Update current department if we have a new one
        if dept_cell and 'Total' not in dept_cell:
            current_dept = dept_cell

        # Skip "Total" rows and invalid entries
        if 'Total' in team_name_cell or team_name_cell in ['#N/A', '']:
            # If team name is empty but we have institution, use current_team
            if team_name_cell == '' and institution_raw and current_team:
                team_name = current_team
            else:
                # Reset current team on total rows or skip
                if 'Total' in team_name_cell:
                    current_team = None
                continue
        else:
            # New team name found
            team_name = team_name_cell
            current_team = team_name

            # Track team under its department
            if current_dept:
                if current_dept not in dept_teams:
                    dept_teams[current_dept] = []
                if team_name not in dept_teams[current_dept]:
                    dept_teams[current_dept].append(team_name)

        if not institution_raw:
            continue

        try:
            fte = float(fte_str)
        except ValueError:
            print(f"Warning: Could not parse FTE value '{fte_str}' for team '{team_name}'")
            continue

        # Normalize institution - case-insensitive match
        institution_lower = institution_raw.lower()
        if institution_lower in inst_lookup:
            institution = inst_lookup[institution_lower]
        else:
            institution = 'Other'

        # Initialize team if not seen
        if team_name not in team_data:
            team_data[team_name] = {inst: 0.0 for inst in institutions}
            team_data[team_name]['Other'] = 0.0

        # Add FTE to the appropriate institution
        team_data[team_name][institution] += fte

    return team_data, dept_teams


def output_team_report(team_data, dept_teams, institutions):
    """
    Output FTE totals per team by institution, grouped by department with subtotals.

    Args:
        team_data: dict {team_name: {institution: FTE_total, ...}, ...}
        dept_teams: dict {department: [team_name, ...], ...}
        institutions: List of institution names
    """
    # Build header
    all_cols = institutions + ['Other', 'Total']
    header = f"{'Team':<40} " + " ".join(f"{inst:>10}" for inst in all_cols)
    sep_line = "=" * len(header)
    dash_line = "-" * len(header)

    print("\n" + sep_line)
    print("Team FTE Summary by Institution")
    print(sep_line)
    print(header)
    print(dash_line)

    # Grand totals
    grand_totals = {inst: 0.0 for inst in institutions}
    grand_totals['Other'] = 0.0
    grand_total = 0.0

    # Process each department
    for dept in dept_teams:
        teams = dept_teams[dept]

        # Department totals
        dept_totals = {inst: 0.0 for inst in institutions}
        dept_totals['Other'] = 0.0
        dept_total = 0.0

        # Print department header
        print(f"\n  {dept}")
        print(f"  {'-' * (len(header) - 2)}")

        # Print each team in this department
        for team_name in teams:
            if team_name not in team_data:
                continue
            data = team_data[team_name]
            row_total = sum(data.values())
            dept_total += row_total

            # Build row
            values = []
            for inst in institutions:
                val = data.get(inst, 0.0)
                dept_totals[inst] += val
                values.append(f"{val:>10.2f}")
            other_val = data.get('Other', 0.0)
            dept_totals['Other'] += other_val
            values.append(f"{other_val:>10.2f}")
            values.append(f"{row_total:>10.2f}")

            print(f"    {team_name:<36} " + " ".join(values))

        # Print department subtotal
        dept_row = []
        for inst in institutions:
            grand_totals[inst] += dept_totals[inst]
            dept_row.append(f"{dept_totals[inst]:>10.2f}")
        grand_totals['Other'] += dept_totals['Other']
        dept_row.append(f"{dept_totals['Other']:>10.2f}")
        grand_total += dept_total
        dept_row.append(f"{dept_total:>10.2f}")

        print(f"  {'-' * (len(header) - 2)}")
        print(f"  {dept + ' Total':<38} " + " ".join(dept_row))

    # Print grand totals
    print("\n" + sep_line)
    totals_row = []
    for inst in institutions:
        totals_row.append(f"{grand_totals[inst]:>10.2f}")
    totals_row.append(f"{grand_totals['Other']:>10.2f}")
    totals_row.append(f"{grand_total:>10.2f}")
    print(f"{'GRAND TOTAL':<40} " + " ".join(totals_row))
    print(sep_line + "\n")


def make_team_report(sheet_id, team_sheet, institutions):
    """
    Fetch team sheet and generate FTE report.

    Args:
        sheet_id: Google Sheet ID
        team_sheet: Name/range of the team sheet
        institutions: List of institutions to track
    """
    print(f"Fetching team data from sheet: {team_sheet}")
    result = get_sheet(sheet_id, team_sheet)
    values = result.get('values', [])

    if not values:
        print("No data found in team sheet.")
        return

    team_data, dept_teams = process_team_sheet(values, institutions)
    output_team_report(team_data, dept_teams, institutions)


# MAIN

parser = argparse.ArgumentParser()
parser.add_argument('id', help="""ID of the google sheet like
                               18wu9f4ov79YDMR1CTEciqAhCawJ7n47C8L9pTAxe""")
parser.add_argument('sheets', nargs='*',
                    help="""Sheet names  and ranges to process
                            within the google sheet e.g. Model!A1:H""")
parser.add_argument("--depth", help="make tree pdf stopping at depth ", type=int, default=100)
parser.add_argument("--land", help="make tree pdf landscape rather than portrait default portrait (1 to make full landscape, 2 mixed)", type=int, default=None)
parser.add_argument("--team", help="Process team sheet and output FTE summary by institution", action="store_true")
parser.add_argument("--team-sheet", help="Name of the team sheet (default: 'team')", default="team")
parser.add_argument("--tree-sheet", help="Name of the tree sheet to process (e.g. 'Model!A1:H')")
parser.add_argument("--institutions", help="Comma-separated list of institutions to track (default: SLAC,IN2P3,UK,AURA,UW,Princeton)",
                    default=",".join(DEFAULT_INSTITUTIONS))
args = parser.parse_args()
outdepth = args.depth
land = args.land

# Parse institutions list
institutions = [inst.strip() for inst in args.institutions.split(',')]

team_data = None
dept_teams = None

# Process team data if --team is specified
if args.team:
    print(f"Fetching team data from sheet: {args.team_sheet}")
    team_result = get_sheet(args.id, args.team_sheet)
    team_values = team_result.get('values', [])
    if team_values:
        team_data, dept_teams = process_team_sheet(team_values, institutions)
        # If only --team (no tree sheet), output the FTE report
        if not args.tree_sheet and not args.sheets:
            output_team_report(team_data, dept_teams, institutions)

# Process tree if --tree-sheet or positional sheets are specified
if args.tree_sheet or args.sheets:
    if land is None:
        print('Output portrait')
    else:
        print('Output landscape ', land)

    sheetId = args.id
    
    # Use --tree-sheet if specified, otherwise use positional sheets
    if args.tree_sheet:
        sheets = [args.tree_sheet]
    else:
        sheets = args.sheets
    
    for r in sheets:
        print("Google %s , Sheet %s" % (sheetId, r))
        result = get_sheet(sheetId, r)
        values = result.get('values', [])
        makeTree(values, team_data, institutions if team_data else None)
elif not args.team:
    parser.error("Either --tree-sheet, sheets argument, or --team is required")

