import * as msal from "@azure/msal-node";
import debug from "debug";
import { getAuthenticatedClient } from "./graph";

const log = debug("app:skills");

const importSkills = async (data: string, msalClient: msal.ConfidentialClientApplication, homeAccountId: string) => {
    
    const client = getAuthenticatedClient(msalClient, homeAccountId, true); // use app permissions

    const lines = data.split("\n");

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line.length > 0) {
            const [upn, skill, proficiency, collaborationTags, allowedAudiences] = line.split(";");
            log(`Checking skill ${skill} for ${upn}`);

            try {
                const request = client
                    .api(`/users/${upn.trim()}/profile/skills`)
                    .version("beta")
                    .filter(`displayName eq '${skill.trim()}' and source/type/any(s:s ne 'UPA')`); // remove all UPA properties
                const currentSkills = await request.get();

                const body = {
                    displayName: skill.trim(),
                    allowedAudiences: allowedAudiences.trim(),
                    collaborationTags: collaborationTags.trim().split(","),
                    proficiency: proficiency.trim(), // generalProfessional does not work
                    source: { type: ["Skillsync"] } // invalid syntax in documentation
                };

                if (currentSkills.value.length > 0) {
                    log(`Updating skill ${skill} for ${upn}`);
                    const update = await client
                        .api(`/users/${upn.trim()}/profile/skills/${currentSkills.value[0].id}`)
                        .version("beta")
                        .patch(body);
                } else {
                    log(`Adding skill ${skill} for ${upn}`);
                    const update = await client
                        .api(`/users/${upn.trim()}/profile/skills`)
                        .version("beta")
                        .post(body);
                    log(`Added skill ${skill} for ${upn} with result ${update.id}`);
                };

            } catch (err) {
                log(err);
            }

        }
    }

}
export default importSkills;