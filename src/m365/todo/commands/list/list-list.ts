import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import DelegatedGraphCommand from '../../../base/DelegatedGraphCommand';
import commands from '../../commands';
import { ToDoList } from '../../ToDoList';

class TodoListListCommand extends DelegatedGraphCommand {
  public get name(): string {
    return commands.LIST_LIST;
  }

  public get description(): string {
    return 'Returns a list of Microsoft To Do task lists';
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'id'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const items: any = await odata.getAllItems<ToDoList>(`${this.resource}/v1.0/me/todo/lists`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TodoListListCommand();