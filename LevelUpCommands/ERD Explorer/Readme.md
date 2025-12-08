# ERD Explorer - Why 2 versions

For the ERD Explorer I have decided to provide 2 different versions. As these are provided for inspiration maybe someone in the community has a good idea how to solve this.

When importing JS code inside of LevelUp some functionalities are highlighted as a warning - e.g., setTimeout, innerHtml, or setInterval.

In the version where I am recording every table, I am using the setInterval to capture navigation from one record to a related record e.g., when clicking on a lookup or subgrid.

## Why LevelUp Shows a Warning</br>

The LevelUp extension/validator likely flags setInterval because:

Concern - Reason</br>
Memory leak risk: If the interval isn't cleared when overlay closes, it keeps running </br>
Performance: Continuous polling every 500ms consumes resources </br>
Best practice: Event-driven approaches are preferred over polling </br>

### The Code Does Handle Cleanup</br>

Looking at both files, they do properly clear the interval:

In stopRecording(): clearInterval(recordingIntervalId)
In closeOverlay(): clearInterval(recordingIntervalId)
So the warning is likely just a general lint/code quality warning, not an actual bug. The cleanup is implemented correctly.

If you want to avoid the warning entirely, use the ERDExplorerV1.4limited.js file (which has no setInterval), but remember it won't auto-detect subgrid/lookup navigation - you have to click manually "capture" to add the table to the selection
