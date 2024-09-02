import { AdministrativeUnit } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import aadCommands from '../../aadCommands';

class EntraAdministrativeUnitListCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of administrative units';
  }

  public alias(): string[] | undefined {
    return [aadCommands.ADMINISTRATIVEUNIT_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'visibility'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.ADMINISTRATIVEUNIT_LIST, commands.ADMINISTRATIVEUNIT_LIST);

    try {
      const results = await odata.getAllItems<AdministrativeUnit>(`${this.resource}/v1.0/directory/administrativeUnits`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraAdministrativeUnitListCommand();