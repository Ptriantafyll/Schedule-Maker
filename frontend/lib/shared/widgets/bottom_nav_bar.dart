import 'package:flutter/material.dart';

class BottomNavBar extends StatelessWidget {
  const BottomNavBar({
    super.key,
    required this.currentIndex,
    required this.showAdminTab,
    required this.onTabSelected,
  });

  final int currentIndex;
  final bool showAdminTab;
  final ValueChanged<int> onTabSelected;

  @override
  Widget build(BuildContext context) {
    return NavigationBar(
      selectedIndex: currentIndex,
      onDestinationSelected: onTabSelected,
      destinations: [
        NavigationDestination(
          icon: const Icon(Icons.grid_view),
          label: 'Dashboard',
        ),
        NavigationDestination(
          icon: const Icon(Icons.calendar_month),
          label: 'Schedule',
        ),
        NavigationDestination(
          icon: const Icon(Icons.swap_horiz),
          label: 'Requests',
        ),
        if (showAdminTab)
          NavigationDestination(
            icon: const Icon(Icons.admin_panel_settings),
            label: 'Admin',
          ),
      ],
    );
  }
}
