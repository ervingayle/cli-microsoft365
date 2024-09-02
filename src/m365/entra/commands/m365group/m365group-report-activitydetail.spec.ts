import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./m365group-report-activitydetail');
import aadCommands from '../../aadCommands';
import Command from '../../../../Command';

describe(commands.M365GROUP_REPORT_ACTIVITYDETAIL, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.M365GROUP_REPORT_ACTIVITYDETAIL);
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
    assert.deepStrictEqual(alias, [aadCommands.M365GROUP_REPORT_ACTIVITYDETAIL]);
  });

  it('gets details about Microsoft 365 Groups activity by group for the given period', async () => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')`) {
        return `
        Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period
        2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7
        2019-10-01,D.Marketing,True,user2@sharepointrider.onmicrosoft.com,2019-05-30,Private,4,0,,,,,,413,3882649,4,1596856,02826124-adbe-4d57-8ccb-a2b5647cad14,7
        `;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { period: 'D7' } });
    assert.strictEqual(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')");
    assert.strictEqual(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
  });
});
