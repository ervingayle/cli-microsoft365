import { Group } from '@microsoft/microsoft-graph-types';
import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { cli } from '../../../../cli/cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./m365group-list');
import aadCommands from '../../aadCommands';

describe(commands.M365GROUP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.M365GROUP_LIST]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'mailNickname', 'siteUrl']);
  });

  it('lists Microsoft 365 Groups in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('lists Microsoft 365 Groups in the tenant (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('lists Microsoft 365 Groups without owners in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$expand=owners&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private",
              "owners": []
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private",
              "owners": [{
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff3",
                "displayName": "Joseph Velliah"
              },
              {
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff4",
                "displayName": "Bose Velliah"
              }]
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { orphaned: true } });
    assert([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "displayName": "Team 1",
        "mailNickname": "team_1"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "displayName": "Team 2",
        "mailNickname": "team_2"
      }
    ]);
  });

  it('lists Microsoft 365 Groups without owners in the tenant (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$expand=owners&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private",
              "owners": []
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private",
              "owners": [{
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff3",
                "displayName": "Joseph Velliah"
              },
              {
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff4",
                "displayName": "Bose Velliah"
              }]
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, orphaned: true } });
    assert([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "displayName": "Team 1",
        "mailNickname": "team_1"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "displayName": "Team 2",
        "mailNickname": "team_2"
      }
    ]);
  });

  it('lists Microsoft 365 Groups filtering on displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Team' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('lists Microsoft 365 Groups filtering on mailNickname', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'team')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { mailNickname: 'team' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('lists Microsoft 365 Groups filtering on displayName and mailNickname', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team') and startswith(MailNickname,'team')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Team', mailNickname: 'team' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('escapes special characters in the displayName filter', async () => {
    const displayName = 'Team\'s #';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'${formatting.encodeQueryParameter(displayName)}')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team's #1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team's #2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: displayName } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team's #1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team's #2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('escapes special characters in the mailNickname filter', async () => {
    const mailNickName = 'team\'s #';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'${formatting.encodeQueryParameter(mailNickName)}')&$top=100`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { mailNickname: mailNickName } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('lists Microsoft 365 Groups in the tenant served in pages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27",
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27`) {
        return {
          "value": [
            {
              "id": "310d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 3",
              "displayName": "Team 3",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_3",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "4157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 4",
              "displayName": "Team 4",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_4",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      },

      {
        "id": "310d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 3",
        "displayName": "Team 3",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_3",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "4157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 4",
        "displayName": "Team 4",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_4",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('handles error when retrieving second page of Microsoft 365 Groups', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27",
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27`) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred'));
  });

  it('lists all properties for output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private"
      }
    ]));
  });

  it('include site URLs of Microsoft 365 Groups', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { includeSiteUrl: true } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private",
        "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
      },
      {
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private",
        "siteUrl": "https://contoso.sharepoint.com/sites/team_2"
      }
    ]));
  });

  it('include site URLs of Microsoft 365 Groups (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, includeSiteUrl: true } });
    assert(loggerLogSpy.calledWith([
      <Group>{
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private",
        "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
      },
      <Group>{
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private",
        "siteUrl": "https://contoso.sharepoint.com/sites/team_2"
      }
    ]));
  });

  it('include site URLs of Microsoft 365 Groups. one group without site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "value": [
            <Group>{
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            <Group>{
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return { webUrl: "" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { includeSiteUrl: true } });
    assert(loggerLogSpy.calledWith([
      <Group>{
        "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-07T13:58:01Z",
        "description": "Team 1",
        "displayName": "Team 1",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_1@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_1",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_1@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-07T13:58:01Z",
        "securityEnabled": false,
        "visibility": "Private",
        "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
      },
      <Group>{
        "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2017-12-17T13:30:42Z",
        "description": "Team 2",
        "displayName": "Team 2",
        "groupTypes": [
          "Unified"
        ],
        "mail": "team_2@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "team_2",
        "onPremisesLastSyncDateTime": null,
        "onPremisesProvisioningErrors": [],
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "proxyAddresses": [
          "SMTP:team_2@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2017-12-17T13:30:42Z",
        "securityEnabled": false,
        "visibility": "Private",
        "siteUrl": ""
      }
    ]));
  });

  it('handles error when retrieving Microsoft 365 Group url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
          "value": [
            <Group>{
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            <Group>{
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        throw 'An error has occurred';
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { includeSiteUrl: true } } as any), new CommandError('An error has occurred'));
  });

  it('passes validation if only includeSiteUrl option set', async () => {
    const actual = await command.validate({ options: { includeSiteUrl: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only orphaned option set', async () => {
    const actual = await command.validate({ options: { orphaned: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if no options set', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
