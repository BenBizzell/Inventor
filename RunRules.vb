ThisDoc.Launch("C:\Users\240126022\Documents\Inventor\Flow Rack\CONFIG RACK.iam")
iLogicVb.RunRule("CreateRack")
RuleParametersOutput()

ThisDoc.Launch("C:\Users\240126022\Documents\Inventor\Flow Rack\RACK_SHELF.iam")
iLogicVb.RunRule("RACK_SHELF:1", "UpdateRoller")
RuleParametersOutput()

ThisDoc.Launch("C:\Users\240126022\Documents\Inventor\Flow Rack\CONFIG RACK.iam")
iLogicVb.UpdateWhenDone = True
ThisDoc.Save