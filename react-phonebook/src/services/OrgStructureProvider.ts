import { IOrgUnit, OrgUnit } from "../classes/IOrgUnit";

export class OrgStructureProvider {
    private static dummyData: OrgUnit[] = [
                {
                    title: "Department 1",
                    child_orgunits: [
                        {
                            title: "Block 1",
                            child_orgunits: [
                                {
                                    title: "Staff1",
                                    child_orgunits: []
                                },
                                {
                                    title: "Staff2",
                                    child_orgunits: [
                                        {
                                            title: "Staff2-1",
                                            child_orgunits: []
                                        }
                                    ]
                                },
                                {
                                    title: "Staff3",
                                    child_orgunits: []
                                }
                            ]
                        },
                        {
                            title: "Block 2",
                            child_orgunits: []
                        },
                        {
                            title: "Block 3",
                            child_orgunits: []
                        }
                    ]
                },
                {
                    title: "Department 2",
                    child_orgunits: []
                },
                {
                    title: "Department 3",
                    child_orgunits: []
                },
                {
                    title: "Department 4",
                    child_orgunits: []
                }
            
        
    ];

    public static async getAllTree():Promise<IOrgUnit[]> {
        return new Promise<IOrgUnit[]>(
            (
              resolve: (units:IOrgUnit[]) => void,
              reject: (error:any) => void
            ) => {
              window.setTimeout(() => {
                resolve(
                    this.dummyData
                );
              }, 2000)
            }
          );
    } 
}