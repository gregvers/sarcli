#!/usr/bin/env python3

import xlrd
import json
import sys
import argparse
import os.path
import re
import collections

class SARconfig:
    def __init__(self, sar_file, partslist_file):
        self.config = {}    # master config dictionary that contains other dictionaries
        self.config["customer"] = {}       # sub-dictionary containing customer and site info
        self.config["subscriptions"] = {}  # sub-dictionary containing customer's subscriptions
        self.config["hw"] = collections.OrderedDict()             # sub-dictionary containing hardware config for the different racks
        self.config["network"] = {}        # sub-dictionary containing network configuration
        self.parts = collections.OrderedDict()     # dictionary with the parts sku and descriptions
        self.bom = collections.OrderedDict()       # dictionary for the bill of material for a given config
        wb = xlrd.open_workbook(sar_file)
        self.sheet_revision = wb.sheet_by_name('Revision History')
        self.sheet_contact = wb.sheet_by_name('Contact Information')
        self.sheet_customersite = wb.sheet_by_name('Customer and Site Requirements')
        self.sheet_subscriptions = wb.sheet_by_name('NEW Cloud Subscriptions')
        self.sheet_hwrequirements = wb.sheet_by_name('NEW Hardware Requirements')
        self.sheet_orderinformation = wb.sheet_by_name('NEW Order Information')
        self.config["sar release"] = int(self.sheet_revision.cell_value(0, 20))  # Revision History sheet, cell U1
        self.parts = self.load_parts(partslist_file)   # populate the dictionary of parts
        self.load_customer_info()      # populate the dictionary of customer and site info
        self.load_subscription_info()  # populate the dictionary of customer's subscriptions
        self.load_rack_info()          # populate the dictionary of hardware config for the different racks
        self.bom = self.build_bom()    # populate the dictionary for the bill of material

    def locate_data_in_sar_file(self):
        row = {}
        col = {}
        if self.config["sar release"] in (20180516, 20180515, 20180514, 20180511, 20180701, 20180725):
            # Contact sheet
            col["customer"] = 5
            row["customer name"] = 6
            # Customer and site sheet
            col["install location"] = 5
            row["country"] = 29
            row["indirect sale"] = 31
            # Cloud subscriptions sheet
            col["subscription"] = 7 # H
            row["occ_cp_subscription"] = 10  # H11
            row["occ_compute_subscription"] = 12  # H13
            row["occ_blockstorage_subscription"] = 14  # H15
            row["occ_blockstoragehighio_subscription"] = 16  # H17
            row["occ_objectstorage_subscription"] = 18  # H19
            row["exacc_base_subscription"] = 25  # H26
            row["exacc_quarter_subscription"] = 27  # H28
            row["exacc_half_subscription"] = 29  # H30
            row["exacc_full_subscription"] = 31  # H32
            row["bdcc_starter_subscription"] = 38  # H39
            row["bdcc_additional_subscription"] = 40  # H41
            col["rack_deployed"] = 63  # BL8
            row["rack_deployed"] = 7  # BL8
            col["occ_rack_allocation"] = 44  # AS8
            col["exacc_rack_allocation"] = 56  # BE8
            col["bdcc_rack_allocation"] = 60  # BI8
            col["tor_deployed"] = 68  # BQ8
            col["spine_deployed"] = 69  # BR8
            # Hardware requirements sheet
            row["pdu"] = 21
            col["pdu_type"] = 4
            col["pdu_whip_count"] = 10
            row["upstream_cable"] = 21
            col["upstream_cable_type"] = 16
            col["upstream_cable_length"] = 17
            col["upstream_cable_count"] = 18
            row["rack_connected"] = 21
            col["rack_connected"] = 13
            col["distance"] = 15
            row["distance"] = 21
            # Order Information
            row["bom"] = 8
            col["bom"] = 2
        else:
            sys.exit("Unsupported SAR release")
        return (row,col)

    def load_customer_info(self):
        (row, col) = self.locate_data_in_sar_file()
        self.config["customer"]["name"] = self.sheet_contact.cell(row["customer name"], col["customer"]).value
        country = self.sheet_customersite.cell(row["country"], col["install location"]).value
        if country[-4:] == "(**)":
            self.config["customer"]["country"] = country[:-4]
            if self.sheet_customersite.cell(row["indirect sale"], col["install location"]).value == "Yes":
                self.config["customer"]["indirect"] = True
            else:
                self.config["customer"]["indirect"] = False
        else:
            self.config["customer"]["country"] = country
            self.config["customer"]["indirect"] = False

    def load_subscription_info(self):
        (row, col) = self.locate_data_in_sar_file()
        self.config["subscriptions"]["OCC CP"] = int(
            self.sheet_subscriptions.cell(row["occ_cp_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["occ_cp_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["OCC Compute"] = int(
            self.sheet_subscriptions.cell(row["occ_compute_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["occ_compute_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["OCC Block storage"] = int(
            self.sheet_subscriptions.cell(row["occ_blockstorage_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["occ_blockstorage_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["OCC Block storage High I/O"] = int(
            self.sheet_subscriptions.cell(row["occ_blockstoragehighio_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["occ_blockstoragehighio_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["OCC Object storage"] = int(
            self.sheet_subscriptions.cell(row["occ_objectstorage_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["occ_objectstorage_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["ExaCC X7 - Base System"] = int(
            self.sheet_subscriptions.cell(row["exacc_base_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["exacc_base_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["ExaCC X7 - Quarter System"] = int(
            self.sheet_subscriptions.cell(row["exacc_quarter_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["exacc_quarter_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["ExaCC X7 - Half System"] = int(
            self.sheet_subscriptions.cell(row["exacc_half_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["exacc_half_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["ExaCC X7 - Full System"] = int(
            self.sheet_subscriptions.cell(row["exacc_full_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["exacc_full_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["BDCC - Starter Pack - 3 Nodes"] = int(
            self.sheet_subscriptions.cell(row["bdcc_starter_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["bdcc_starter_subscription"], col["subscription"]).value != '' else 0
        self.config["subscriptions"]["BDCC - Additional Nodes"] = int(
            self.sheet_subscriptions.cell(row["bdcc_additional_subscription"], col["subscription"]).value) \
                if self.sheet_subscriptions.cell(row["bdcc_additional_subscription"], col["subscription"]).value != '' else 0

    def load_parts(self, partslist_file):
        with open(partslist_file) as f:
            parts = json.load(f, object_pairs_hook=collections.OrderedDict)
        return parts

    def get_rack_count(self, row, col):
        rack_deployed = self.sheet_subscriptions.col_slice(colx=col["rack_deployed"], start_rowx=row["rack_deployed"],
                                                     end_rowx=row["rack_deployed"] + 11)
        rackcount = 0
        for cell in rack_deployed:
            if int(cell.value) != 0:
                rackcount = rackcount + 1
            else:
                break
        return rackcount

    def load_rack_info(self):
        (row, col) = self.locate_data_in_sar_file()
        rackcount = self.get_rack_count(row, col)
        for rackid in range(rackcount):
            rackname = "rack" + str(rackid + 1)
            self.config["hw"][rackname] = {}
            self.config["hw"][rackname]["id"] = rackid + 1
            if self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"]).value > 0:
                if self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"] + 1).value == 1:
                    self.config["hw"][rackname]["type"] = "OCC CP"
                    self.config["hw"][rackname]["CP qty"] = int(
                        self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"] + 1).value)
                else:
                    self.config["hw"][rackname]["type"] = "OCC"
                self.config["hw"][rackname]["block hdd qty"] = int(
                    self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"] + 2).value)
                self.config["hw"][rackname]["object qty"] = int(
                    self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"] + 3).value)
                self.config["hw"][rackname]["oasg qty"] = int(
                    self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"] + 4).value)
                self.config["hw"][rackname]["block ssd qty"] = int(
                    self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"] + 6).value)
                self.config["hw"][rackname]["compute qty"] = int(
                    self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["occ_rack_allocation"] + 7).value)
            elif self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["exacc_rack_allocation"]).value > 0:
                self.config["hw"][rackname]["type"] = "ExaCC Full"
            elif self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["exacc_rack_allocation"] + 1).value > 0:
                self.config["hw"][rackname]["type"] = "ExaCC Half"
            elif self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["exacc_rack_allocation"] + 2).value > 0:
                self.config["hw"][rackname]["type"] = "ExaCC Quarter"
            elif self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["exacc_rack_allocation"] + 3).value > 0:
                self.config["hw"][rackname]["type"] = "ExaCC Base"
            elif self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["bdcc_rack_allocation"]).value > 0:
                self.config["hw"][rackname]["type"] = "BDCC Full"
                self.config["hw"][rackname]["node qty"] = int(self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["bdcc_rack_allocation"] + 2).value)
            elif self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["bdcc_rack_allocation"] + 1).value > 0:
                self.config["hw"][rackname]["type"] = "BDCC Starter"
                self.config["hw"][rackname]["node qty"] = int(self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["bdcc_rack_allocation"] + 2).value)
            self.config["hw"][rackname]["ToR deployed"] = True \
                if self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["tor_deployed"]).value == 'Y' else False
            self.config["hw"][rackname]["spine deployed"] = True \
                if self.sheet_subscriptions.cell(row["rack_deployed"] + rackid, col["spine_deployed"]).value == 'Y' else False
            self.config["hw"][rackname]["pdu type"] = self.sheet_hwrequirements.cell(row["pdu"] + rackid, col["pdu_type"]).value
            self.config["hw"][rackname]["internal connection"] = {}
            if self.config["hw"][rackname]["ToR deployed"] or self.config["hw"][rackname]["type"] in ("BDCC Full", "BDCC Starter", "BDCC Addn"):
                self.config["hw"][rackname]["upstream cable type"] = self.sheet_hwrequirements.cell(row["upstream_cable"] + rackid, col["upstream_cable_type"]).value
                self.config["hw"][rackname]["upstream cable length"] = self.sheet_hwrequirements.cell(row["upstream_cable"] + rackid, col["upstream_cable_length"]).value
                self.config["hw"][rackname]["upstream cable count"] = int(self.sheet_hwrequirements.cell(row["upstream_cable"] + rackid, col["upstream_cable_count"]).value) \
                    if self.sheet_hwrequirements.cell(row["upstream_cable"] + rackid, col["upstream_cable_count"]).value != "" else 0
                if self.config["hw"]["rack1"]["spine deployed"]:
                    self.config["hw"][rackname]["internal connection"]["type"] = "ToR to spine"
                    if rackid == 0:
                        self.config["hw"][rackname]["internal connection"]["distance"] = self.sheet_hwrequirements.cell(
                            row["distance"] + rackid + 1, col["distance"]).value  # distance for Rack2 instead
                    else:
                        self.config["hw"][rackname]["internal connection"]["distance"] = self.sheet_hwrequirements.cell(
                            row["distance"] + rackid, col["distance"]).value
                        self.config["hw"][rackname]["internal connection"]["distance to OOB"] = \
                            self.config["hw"][rackname]["internal connection"]["distance"]
                else:
                    if rackid > 0:
                        self.config["hw"][rackname]["internal connection"]["type"] = "ToR to ToR"
                        self.config["hw"][rackname]["internal connection"]["distance"] = self.sheet_hwrequirements.cell(
                            row["distance"] + rackid, col["distance"]).value
                        self.config["hw"][rackname]["internal connection"]["distance to OOB"] = \
                            self.config["hw"][rackname]["internal connection"]["distance"]
            else:
                self.config["hw"][rackname]["internal connection"]["type"] = "eth to ToR"
                self.config["hw"][rackname]["internal connection"]["to rack"] = self.sheet_hwrequirements.cell(
                    row["rack_connected"] + rackid, col["rack_connected"]).value
                self.config["hw"][rackname]["internal connection"]["distance"] = self.sheet_hwrequirements.cell(
                    row["distance"] + rackid, col["distance"]).value
                self.config["hw"][rackname]["internal connection"]["distance to OOB"] = self.config["hw"][rackname]["internal connection"]["distance"]   # gap: not captured in SAR

    def init_rack_partsqty(self):
        rackpartsqty = collections.OrderedDict()
        for partname, info in self.parts.items():
            rackpartsqty[partname] = 0
        return rackpartsqty

    def build_bom(self):
        partsqty = collections.OrderedDict()   # dictionary of racks associated with a sub-dictionary of parts/qty to include in the BOM
        for rackname, rackconfig in self.config["hw"].items():
            partsqty[rackname] = self.init_rack_partsqty()   # sub-dictionary of parts/qty to include in the BOM
            # adding parts for eack type of rack
            if rackconfig["type"] in ("OCC CP", "OCC"):     # adding OCC parts
                partsqty[rackname]["OCC"] = 1
                if self.config["customer"]["indirect"]:
                    partsqty[rackname]["OCC rack resale"] = 1
                else:
                    partsqty[rackname]["OCC rack"] = 1
                if rackconfig["type"] == "OCC CP":      # adding OCC Control Plane parts
                    if rackconfig["CP qty"] == 1:       # OCC CP contains 5 admin nodes, 2 regular compute and 1 block SSD
                        partsqty[rackname]["OCC admin"] += 5
                        partsqty[rackname]["OCC compute"] += 2
                        partsqty[rackname]["OCC block ssd"] += 1
                partsqty[rackname]["OCC compute"] += rackconfig["compute qty"]
                partsqty[rackname]["OCC block ssd"] += rackconfig["block ssd qty"]
                partsqty[rackname]["OCC block hdd"] += rackconfig["block hdd qty"]
                partsqty[rackname]["OCC object"] += rackconfig["object qty"]
                partsqty[rackname]["OCC software"] += 1
                partsqty[rackname]["Juniper support ToR"] += 1
                partsqty[rackname]["Juniper support Spine"] += 1
                partsqty[rackname]["install service Engineered Systems"] += 1
                partsqty[rackname]["Cisco support"] += 1
                partsqty[rackname]["OASG"] += rackconfig["oasg qty"]
                if partsqty[rackname]["OASG"] > 0:    # OASG requires 2 power cables
                    partsqty[rackname]["power cable"] += 2
                partsqty[rackname]["install service OASG"] += rackconfig["oasg qty"]   # install service to include with OASG
            elif rackconfig["type"] in ("ExaCC Base", "ExaCC Quarter", "ExaCC Half", "ExaCC Full"):   # adding ExaCC parts
                if rackconfig["type"] == "ExaCC Base":
                    partsqty[rackname]["ExaCC"] = 1
                elif rackconfig["type"] == "ExaCC Quarter":
                    partsqty[rackname]["ExaCC"] = 1
                elif rackconfig["type"] == "ExaCC Half":
                    partsqty[rackname]["ExaCC"] = 1
                elif rackconfig["type"] == "ExaCC Full":
                    partsqty[rackname]["ExaCC"] = 1
                if rackconfig["ToR deployed"]:
                    partsqty[rackname]["ExaCC ToR"] = 2
                    partsqty[rackname]["ExaCC cable kit"] = 1
            elif rackconfig["type"] in ("BDCC Full", "BDCC Starter"):               # adding BDCC parts
                partsqty[rackname]["BDCC"] = 1
                partsqty[rackname]["BDCC base rack"] = 1
                if rackconfig["type"] == "BDCC Starter":
                    partsqty[rackname]["BDCC starter rack"] = 1
                elif rackconfig["type"] == "BDCC Full":
                    partsqty[rackname]["BDCC full rack"] = 1
                partsqty[rackname]["BDCC node"] = rackconfig["node qty"]
            # adding North-South cables and transceivers
            if rackconfig["ToR deployed"]:           # we provide upstream cables and transceivers for ToR switches
                cablename = "cable " + rackconfig["upstream cable type"] + " " + rackconfig["upstream cable length"]
                partsqty[rackname][cablename] = rackconfig["upstream cable count"]
                if rackconfig["upstream cable type"] == "MPO_4LC":
                    partsqty[rackname]["TRX QSFP+ ESR4"] = rackconfig["upstream cable count"]      # MPO-4LC cables require ESR4 transceivers
                else:
                    partsqty[rackname]["TRX QSFP+ SR4"] = rackconfig["upstream cable count"]      # MPO-MPO cables require SR4 transceivers
            elif rackconfig["type"] in ("BDCC Full", "BDCC Starter"):      # for BDCC, the transceivers are already included as part of the NM2 GW so we just provide upstream cables
                cablename = "cable " + rackconfig["upstream cable type"] + " " + rackconfig["upstream cable length"]
                partsqty[rackname][cablename] = rackconfig["upstream cable count"]
            # adding East-West cables and transceivers
            if rackname == "rack1":
                if rackconfig["spine deployed"]:   # rack1 with a spine includes 4 short MPO_MPO cables to connect to spine, 2 go in rack1, 2 go in rack2
                    partsqty[rackname]["spine"] = 1
                    partsqty[rackname]["cable MPO_MPO 5m"] += 4
                    partsqty[rackname]["TRX QSFP+ SR4"] += 8
            elif rackconfig["ToR deployed"]:   # rack2 with a spine requires 4 MPO_MPO cables to connect to spine, 2 go in rack1, 2 go in rack2
                                               # other racks with ToR require 4 MPO_MPO cables to connect to spine
                cablename = "cable MPO_MPO " + rackconfig["internal connection"]["distance"]
                partsqty[rackname][cablename] += 4
                partsqty[rackname]["TRX QSFP+ SR4"] += 8
            elif rackconfig["type"] in ("BDCC Full", "BDCC Starter", "BDCC Addn"):
                cablename = "cable MPO_MPO " + rackconfig["internal connection"]["distance"]
                partsqty[rackname][cablename] += 2
                partsqty[rackname]["TRX QSFP+ SR4"] += 2
            elif not rackconfig["ToR deployed"]:  # ExaCC without ToR connect to a ExaCC rack using LC cables or to a OCC rack using copper cables
                connected_rack = rackconfig["internal connection"]["to rack"]
                if self.config["hw"][connected_rack]["type"] in ("ExaCC Base", "ExaCC Quarter", "ExaCC Half", "ExaCC Full"):
                    cablename = "cable LC " + rackconfig["internal connection"]["distance"]
                    partsqty[rackname][cablename] += 10
                    partsqty[rackname]["TRX SFP+"] += 8
                elif self.config["hw"][connected_rack]["type"] in ("OCC CP", "OCC"):
                    cablename = "cable CAT6 " + rackconfig["internal connection"]["distance"]
                    partsqty[rackname][cablename] += 9
            # adding copper cables for OOB switch connections
            if rackconfig["type"] == "OCC CP":
                partsqty[rackname]["cable CAT6 5m"] += 1     # cable to connect the OASG to the OOB switch both located in the same rack
            else:
                if rackconfig["internal connection"] != {}:   # exclude case of ExaCC scenario7 where ExaCC rack has no attachment to OCC CP
                    cablename = "cable CAT6 " + rackconfig["internal connection"]["distance"]   # cable to connect the rack OOB switch to the rack1 OOB switch
                    partsqty[rackname][cablename] += 1
            # PDU
            if rackconfig["pdu type"] == "* Single-Phase 2(Two)x22kVA High Voltage Power Supplies (EMEA & APAC (excluding Japan /Taiwan)":
                partsqty[rackname]["PDU 1phase-230V-22kVA"] += 1
            elif rackconfig["pdu type"] == "* Single-Phase 2(Two)x15kVA Low Voltage Power Supplies (Americas / Japan /Taiwan)":
                partsqty[rackname]["PDU 3phase-120V-15kVA"] += 1
            elif rackconfig["pdu type"] == "* Single-Phase 2(Two)x22kVA Low Voltage Power Supplies (Americas / Japan /Taiwan)":
                partsqty[rackname]["PDU 3phase-120V-22kVA"] += 1
            elif rackconfig["pdu type"] == "* Three-Phase 2(Two)x15kVA High Voltage Power Supplies  (EMEA & APAC (excluding Japan /Taiwan)":
                partsqty[rackname]["PDU 3phase-230V-15kVA"] += 1
            elif rackconfig["pdu type"] == "* Three-Phase 2(Two)x24kVA High Voltage Power Supplies  (EMEA & APAC (excluding Japan /Taiwan)":
                partsqty[rackname]["PDU 3phase-230V-24kVA"] += 1
            elif rackconfig["pdu type"] == "* Three-Phase 2(Two)x15kVA Low Voltage Power Supplies (Americas / Japan /Taiwan)":
                partsqty[rackname]["PDU 3phase-120V-15kVA"] += 1
            elif rackconfig["pdu type"] == "* Three-Phase 2(Two)x24kVA Low Voltage Power Supplies (Americas / Japan /Taiwan)":
                partsqty[rackname]["PDU 3phase-120V-24kVA"] += 1
        return partsqty

    def dump_bom(self):
        return json.dumps(self.bom, indent=4)

    def print_bom(self):
        for rack, partslist in self.bom.items():
            print(rack)
            for partnickname, qty in partslist.items():
                if qty > 0:
                    if self.config["hw"][rack]["type"] == "OCC CP" and partnickname == "OCC":
                        print("## ORACLE CLOUD AT CUSTOMER X6 ##")
                        print("Qty    Part  # Description")
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"], self.parts[partnickname]["label"]))
                        print("Type   :New System")
                    elif self.config["hw"][rack]["type"] == "OCC" and partnickname == "OCC":
                        print("## ORACLE CLOUD AT CUSTOMER X6 ##")
                        print("Qty    Part  # Description")
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"],
                                                         self.parts[partnickname]["label"]))
                        print("Type   :Expansion")
                    elif self.config["hw"][rack]["type"] == "ExaCC Base" and partnickname == "ExaCC":
                        print("## EXADATA CLOUD AT CUSTOMER X7 ##")
                        print("Qty    Part  # Description")
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"],
                                                         self.parts[partnickname]["label"]))
                        print("Rack Size : Base Rack")
                    elif self.config["hw"][rack]["type"] == "ExaCC Quarter" and partnickname == "ExaCC":
                        print("## EXADATA CLOUD AT CUSTOMER X7 ##")
                        print("Qty    Part  # Description")
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"],
                                                         self.parts[partnickname]["label"]))
                        print("Rack Size : Quarter Rack")
                    elif self.config["hw"][rack]["type"] == "ExaCC Half" and partnickname == "ExaCC":
                        print("## EXADATA CLOUD AT CUSTOMER X7 ##")
                        print("Qty    Part  # Description")
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"],
                                                         self.parts[partnickname]["label"]))
                        print("Rack Size : Half Rack")
                    elif self.config["hw"][rack]["type"] == "ExaCC Full" and partnickname == "ExaCC":
                        print("## EXADATA CLOUD AT CUSTOMER X7 ##")
                        print("Qty    Part  # Description")
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"],
                                                         self.parts[partnickname]["label"]))
                        print("Rack Size : Full Rack")
                    elif partnickname == "BDCC":
                        print("## BIG DATA CLOUD AT CUSTOMER X7 ##")
                        print("Qty    Part  # Description")
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"],
                                                         self.parts[partnickname]["label"]))
                    else:
                        print('{0:3} x  {1:7}  {2}'.format(str(qty), self.parts[partnickname]["sku"],
                                                         self.parts[partnickname]["label"]))
            print()

    def dump_config(self):
        return json.dumps(self.config, indent=4)

    def dump_partslist(self):
        return json.dumps(self.parts, indent=4)

    def diff_bom(self):
        (row, col) = self.locate_data_in_sar_file()
        sarfilebom = {}
        rackid = 0
        for rack, partslist in self.bom.items():
            bom_fromxls = []
            bom_generated = []
            print(rack + ":")
            sarfilebom[rack] = self.sheet_orderinformation.cell(row["bom"] + rackid, col["bom"]).value
            rack_bom_lines = [y for y in (x.strip() for x in sarfilebom[rack].splitlines()) if y]
            for line in rack_bom_lines:
                m = re.match(r"^(\d+)\s* x \s*(.*?) \s*(.*)$", line)
                if m:
                    if m.group(3)[:2] == "* ":
                        bom_fromxls.append((int(m.group(1)), m.group(2), m.group(3)[2:]))
                    else:
                        bom_fromxls.append((int(m.group(1)), m.group(2), m.group(3)))
                else:
                    print("Ignored line: " + line)
            for partnickname, qty in partslist.items():
                if qty > 0:
                    sku = self.parts[partnickname]["sku"]
                    label = self.parts[partnickname]["label"]
                    bom_generated.append((qty, sku, label))
            #print(bom_fromxls)
            #print(bom_generated)
            for line in bom_fromxls:
                if line in bom_generated:
                    bom_generated.remove(line)
                else:
                    print('in XLS but not generated:    {0:3} x  {1:7}  {2}'.format(str(line[0]), line[1], line[2]))
            for line in bom_generated:
                print('generated but not in XLS:    {0:3} x  {1:7}  {2}'.format(str(line[0]), line[1], line[2]))
            print()
            rackid += 1

######## MAIN #########
def main(argv):
    sar_file = parts_file = ""
    bom = config = False
    parser = argparse.ArgumentParser(prog='sar', usage='%(prog)s command sarfile [options]')
    parser.add_argument("command", type=str, action='store', choices=['config','bom','diff'], help="command")
    parser.add_argument("sarfile", type=str, help="sar file")
    parser.add_argument("-p", "--partsfile", action="store", nargs='?', default="./CatC-partslist.json")
    args = parser.parse_args()
    if not os.path.isfile(args.sarfile):
        print("Specified SAR file is not found: " + args.sarfile)
        sys.exit()
    if not os.path.isfile(args.partsfile):
        print("Parts file is not found: " + args.partsfile)
        sys.exit()
    sar = SARconfig(args.sarfile, args.partsfile)
    if args.command == "config":
        print(sar.dump_config())
    elif args.command == "bom":
        #print(sar.dump_bom())
        sar.print_bom()
    elif args.command == "diff":
        sar.diff_bom()
    else:
        usage()
        sys.exit()


if __name__ == "__main__":
    main(sys.argv[1:])
