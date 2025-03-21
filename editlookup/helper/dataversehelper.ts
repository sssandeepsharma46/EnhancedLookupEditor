export const fetchLookupRecord = async (
    webAPI: ComponentFramework.WebApi,
    lookupId: string,
    lookupSchemaName:string,
    lookupColumnName:string,
    lookupNameColumn: string,
):Promise<{name:string; columnName:string; recordid:string} | null> =>
{
    try{
        //const response = await webAPI.retrieveRecord(lookupSchemaName, lookupId, `?$select=name,${lookupColumnName}`);
        const response = await webAPI.retrieveRecord(
            lookupSchemaName,
            lookupId,
            `?$select=${lookupNameColumn},${lookupColumnName}`
        );

        // Extract values from the response dynamically
        const name = response[lookupNameColumn as keyof typeof response] as string;
        const recordid = response[`${lookupSchemaName}id` as keyof typeof response] as string;
        const columnName = response[lookupColumnName as keyof typeof response] as string;

        // Return the result with dynamic column names

        return{
            name,
            columnName,
            recordid
        };
    } catch(error){
        console.error("Error fetching lookup record:",error);
        return null;
    }
};