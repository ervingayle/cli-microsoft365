import Command from "../Command";
import { CommandOptionInfo } from "./CommandOptionInfo";

export interface CommandInfo {
  aliases: string[] | undefined;
  command: Command;
  defaultProperties: string[] | undefined;
  description: string;
  file: string;
  help?: string;
  name: string;
  options: CommandOptionInfo[];
}