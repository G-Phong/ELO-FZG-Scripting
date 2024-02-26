IXConnection conn = ...;
String objId = ...;


EditInfoZ editZ = new EditInfoZ(EditInfoC.mbMask, SordC.mbLean);
ed = conn.ix.checkoutSord(objId, editZ, LockC.NO);

for (int i = 0; i < ed.mask.lines.length; i++) {
    DocMaskLine dml = ed.mask.lines[i];
    ObjKey okey = ed.sord.objKeys[i];

    // ID is equal to array index: i == dml.id && i == okey.id;

    String okeyValue = "";
    if (okey.data.length != 0) okeyValue = okey.data[0];
    Console.WriteLine(dml.name + "=" + okeyValue);
}


ed.sord.objKeys[2].data; //Inhalt eines Indexwertes