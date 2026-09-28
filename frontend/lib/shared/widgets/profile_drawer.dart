import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';

class ProfileDrawer extends ConsumerWidget {
  const ProfileDrawer({super.key});

  @override
  Widget build(BuildContext context, WidgetRef ref) {
    final currentUser = ref.watch(currentUserProvider);

    return Drawer(
      child: ListView(
        padding: EdgeInsets.zero,
        children: [
          UserAccountsDrawerHeader(
            currentAccountPicture: CircleAvatar(
              child: const Icon(Icons.person, size: 36),
            ),
            accountEmail: Text(currentUser?.email ?? "Not signed in"),
            accountName: Text(currentUser?.fullName ?? "Guest"),
          ),

          if (currentUser != null)
            ListTile(
              leading: const Icon(Icons.badge_outlined),
              title: const Text('Role'),
              subtitle: Text(currentUser.role.displayName),
            ),

          ListTile(
            leading: const Icon(Icons.logout),
            title: const Text('Log Out'),
            onTap: () {
              ref.read(authControllerProvider.notifier).logout();
            },
          ),
        ],
      ),
    );
  }
}
