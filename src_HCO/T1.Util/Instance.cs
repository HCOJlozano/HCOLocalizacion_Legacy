using SAPbouiCOM;

namespace T1.Util
{
    public class Instance
    {
        public static void AddChooseFromList(Form oForm, string objType, string uniqueID)
        {
            try
            {

                ChooseFromListCollection oCFLs = null;
                Conditions oCons = null;
                Condition oCon = null;

                oCFLs = oForm.ChooseFromLists;

                ChooseFromList oCFL = null;
                ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(B1.MainObject.Instance.B1Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = objType;
                oCFLCreationParams.UniqueID = uniqueID;

                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch
            {
                
            }
        }
    }
}
