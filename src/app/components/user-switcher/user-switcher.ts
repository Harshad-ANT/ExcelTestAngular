import { Component, output } from '@angular/core';
import { UserService } from '../../services/user.service';

@Component({
  selector: 'app-user-switcher',
  template: `
    <div class="user-switcher">
      <div class="current-user" (click)="open = !open">
        <span class="avatar" [attr.data-role]="userService.activeUser().role">
          {{ userService.activeUser().name.charAt(0) }}
        </span>
        <span class="user-name">{{ userService.activeUser().name }}</span>
        <span class="role-badge" [attr.data-role]="userService.activeUser().role">
          {{ userService.activeUser().role }}
        </span>
      </div>
      @if (open) {
        <div class="dropdown">
          @for (user of userService.users(); track user.id) {
            <div
              class="user-option"
              [class.active]="user.id === userService.activeUserId()"
              (click)="selectUser(user.id)"
            >
              <span class="avatar small" [attr.data-role]="user.role">{{ user.name.charAt(0) }}</span>
              <span class="option-name">{{ user.name }}</span>
              <span class="role-badge small" [attr.data-role]="user.role">{{ user.role }}</span>
            </div>
          }
        </div>
      }
    </div>
  `,
  styles: [`
    .user-switcher { position: relative; }

    .current-user {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 4px 10px;
      border: 1px solid #ddd;
      border-radius: 8px;
      cursor: pointer;
      background: #fff;
      &:hover { background: #f5f5f5; }
    }

    .avatar {
      width: 28px; height: 28px;
      border-radius: 50%;
      display: flex; align-items: center; justify-content: center;
      font-size: 13px; font-weight: 600; color: #fff;
      &[data-role="admin"] { background: #e53935; }
      &[data-role="editor"] { background: #1e88e5; }
      &[data-role="viewer"] { background: #7cb342; }
      &.small { width: 24px; height: 24px; font-size: 11px; }
    }

    .user-name { font-size: 13px; font-weight: 500; color: #333; }

    .role-badge {
      font-size: 10px; font-weight: 600; text-transform: uppercase;
      padding: 2px 6px; border-radius: 4px;
      &[data-role="admin"] { background: #ffebee; color: #c62828; }
      &[data-role="editor"] { background: #e3f2fd; color: #1565c0; }
      &[data-role="viewer"] { background: #f1f8e9; color: #558b2f; }
      &.small { font-size: 9px; padding: 1px 4px; }
    }

    .dropdown {
      position: absolute; top: 100%; right: 0; margin-top: 4px;
      background: #fff; border: 1px solid #e0e0e0; border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 1000; min-width: 220px; overflow: hidden;
    }

    .user-option {
      display: flex; align-items: center; gap: 8px;
      padding: 10px 14px; cursor: pointer;
      &:hover { background: #f5f5f5; }
      &.active { background: #e8f5e9; }
    }

    .option-name { flex: 1; font-size: 13px; color: #333; }
  `],
})
export class UserSwitcherComponent {
  open = false;
  userChanged = output<string>();

  constructor(public userService: UserService) {}

  selectUser(userId: string) {
    this.userService.switchUser(userId);
    this.open = false;
    this.userChanged.emit(userId);
  }
}
