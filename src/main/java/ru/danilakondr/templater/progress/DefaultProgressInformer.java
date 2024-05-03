/*
 * Copyright (c) 2024 Danila A. Kondratenko
 *
 * This file is a part of UNO Templater.
 *
 * UNO Templater is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * UNO Templater is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with UNO Templater.  If not, see <https://www.gnu.org/licenses/>.
 */

package ru.danilakondr.templater.progress;

public class DefaultProgressInformer implements ProgressInformer {
    private String progressString;
    private boolean silent = false;

    public DefaultProgressInformer(String progressString) {
        this.progressString = progressString;
    }

    public void setProgressString(String progressString) {
        this.progressString = progressString;
    }

    public void setSilent(boolean silent) {
        this.silent = silent;
    }

    @Override
    public void inform(int current, int total) {
        if (silent)
            return;

        if (current == -1 && total == -1)
            System.out.println(progressString + "...");
        else if (total == -1)
            System.out.printf("%s (%d)...%n", progressString, current);
        else
            System.out.printf("%s (%d/%d)...%n", progressString, current, total);
    }
}