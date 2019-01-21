{
  Copyright (C) 2016 by Clever Components

  Author: Sergey Shirokov <admin@clevercomponents.com>

  Website: www.CleverComponents.com

  This file is part of Clever Components Mailer.

  Clever Components Mailer is free software:
  you can redistribute it and/or modify it under the terms of
  the GNU Lesser General Public License version 3
  as published by the Free Software Foundation and appearing in the
  included file COPYING.LESSER.

  Clever Components Mailer is distributed in the hope
  that it will be useful, but WITHOUT ANY WARRANTY; without even the
  implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
  See the GNU Lesser General Public License for more details.

  You should have received a copy of the GNU Lesser General Public License
  along with Clever Components Mailer. If not, see <http://www.gnu.org/licenses/>.

  The current version of Clever Components Mailer needs for
  the non-free library Clever Internet Suite. This is a drawback,
  and we suggest the task of changing
  the program so that it does the same job without the non-free library.
  Anyone who thinks of doing substantial further work on the program,
  first may free it from dependence on the non-free library.
}

program CleverMailer;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  Clever.Mailer in 'Clever.Mailer.pas';

var
  mailer: TMailer;
begin
  try
    mailer := TMailer.Create();
    try
      Writeln(mailer.GetVersionInfo());

      if (ParamCount() < 1) then
      begin
        Writeln('Usage: ', ExtractFileName(ParamStr(0)), ' ConfigFile.ini');
        Exit;
      end;

      mailer.ReadConfiguration(ParamStr(1));

      mailer.SendMessage();
    finally
      mailer.Free();
    end;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
